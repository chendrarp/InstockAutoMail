using MailFactory;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.ServiceProcess;
using System.Threading;
using System.Linq;
using System.Text;
using System.Net;
using InstockAutoMailService.Models;
using NPOI.XSSF.UserModel;
namespace InstockAutoMailService
{
    partial class Instock : ServiceBase
    {
        System.Timers.Timer instockAutoTimer;
        public bool _runFlag;
        public string _DBConnStr, Mail_User, Mail_Password, Mail_Host, _CName;
        private int Mail_Port;
        private double MailTriTime;
        // private double fileMaxSize = 2;
        private static string serviceInstallPath;
        private string strPublicWebIP = string.Empty;
        private Thread _ObaOutputDataThread;
        //private Thread _OqcDisplayFailThread;
        private Thread _PdlOutputDataThread;
        //private Thread _TimeOutDataThread2;
        //private Thread _TimeOutDataThread8;
        //private Thread _TimeOutDataThread24;
        //private Thread _MPKShippedThread;
        //private Thread _VUTVShippedThread;
        //private Thread _VUTVShippedAllThread;
        //private Thread _smsThread;
        //private System.Threading.Thread _instockHourlyOutPutDataThread2;
        //private System.Threading.Thread _mfrAutoOutPutDataThread;
        //private System.Threading.Thread _mfrOutPutDataThread;
        //private System.Threading.Thread _dashBoardInstockFull;

        private System.Threading.Thread _mpkandinstockgapThread;

        ConnectionDapper con = new ConnectionDapper();
        public Instock()
        {
            InitializeComponent();

            //this._DBConnStr = "user id=" + System.Configuration.ConfigurationManager.AppSettings["USER_ID"].ToString() + ";data source=" + System.Configuration.ConfigurationManager.AppSettings["DATA_SOURCE"].ToString() + ";password=" + System.Configuration.ConfigurationManager.AppSettings["PASSWORD"].ToString() + ";max pool size = 512";
            //this.Mail_User = System.Configuration.ConfigurationManager.AppSettings["Mail_User"].ToString();
            //this.Mail_Password = System.Configuration.ConfigurationManager.AppSettings["Mail_Password"].ToString();
            //this.Mail_Host = System.Configuration.ConfigurationManager.AppSettings["Mail_Host"].ToString();
            //this.Mail_Port = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Mail_Port"].ToString());
            //this._CName = System.Configuration.ConfigurationManager.AppSettings["CNAME"].ToString();
            //serviceInstallPath = AppDomain.CurrentDomain.BaseDirectory + "\\";
            //_runFlag = true;
            //SendPublicSMS();
            ////CollectVUTVShippmentData();
            //CollectVUTVAllShippmentData();
            //CollectOBAAutoOutPutData();
            //CollectMFRAutoOutPutData();
            //CollectHourlyInstockOutPutData2();
            //test();
            //CollectPDLAutoOutPutData();
            //CollectMPKShippedAutoOutPutData();
            //CollectDashboardInstockFullData();
            //CollectOQCDisplayFailData();
            CollectMpkInstockGapData();
        }
        protected override void OnStart(string[] args)
        {

            this._DBConnStr = "user id=" + System.Configuration.ConfigurationManager.AppSettings["USER_ID"].ToString() + ";data source=" + System.Configuration.ConfigurationManager.AppSettings["DATA_SOURCE"].ToString() + ";password=" + System.Configuration.ConfigurationManager.AppSettings["PASSWORD"].ToString() + ";max pool size = 512";
            this.Mail_User = System.Configuration.ConfigurationManager.AppSettings["Mail_User"].ToString();
            this.Mail_Password = System.Configuration.ConfigurationManager.AppSettings["Mail_Password"].ToString();
            this.Mail_Host = System.Configuration.ConfigurationManager.AppSettings["Mail_Host"].ToString();
            this.Mail_Port = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Mail_Port"].ToString());
            this._CName = System.Configuration.ConfigurationManager.AppSettings["CNAME"].ToString();
            serviceInstallPath = AppDomain.CurrentDomain.BaseDirectory + "\\";


            if (InstockAutoMail.Default.IsMpkInstockGap)
            {
                WriteLog("IsMpkInstockGap", "OnStart");
                _runFlag = true;
                _mpkandinstockgapThread = new Thread(new ThreadStart(this.CollectMpkInstockGapData));
                _mpkandinstockgapThread.Start();
            }

            //if (InstockAutoMail.Default.IsDashBoardInstockFull)
            //{
            //    WriteLog("IsDashBoardInstockFull", "OnStart");
            //    _runFlag = true;
            //    _dashBoardInstockFull = new Thread(new ThreadStart(this.CollectDashboardInstockFullData));
            //    _dashBoardInstockFull.Start();
            //}
            //if (InstockAutoMail.Default.IsOqcDisplayFail)
            //{
            //    WriteLog("IsOqcDisplayFail", "OnStart");
            //    _runFlag = true;
            //    _OqcDisplayFailThread = new Thread(new ThreadStart(this.CollectOQCDisplayFailData));
            //    _OqcDisplayFailThread.Start();
            //}

            //if (InstockAutoMail.Default.IsAutoTrigger)
            //{
            //    WriteLog("Auto Service Started", " OnStart");
            //    _runFlag = true;
            //    _mfrAutoOutPutDataThread = new Thread(new ThreadStart(this.CollectMFRAutoOutPutData));
            //    _mfrAutoOutPutDataThread.Start();
            //}

            //if (InstockAutoMail.Default.IsManualTrigger)
            //{
            //    WriteLog("Manual Service Started", " OnStart");
            //    _runFlag = true;
            //    string fdate = InstockAutoMail.Default.fromDate;
            //    string tdate = InstockAutoMail.Default.toDate;
            //    _mfrOutPutDataThread = new Thread(new ThreadStart(() => this.CollectMFRManualOutPutData(fdate, tdate)));
            //    _mfrOutPutDataThread.Start();
            //}
            //if (InstockAutoMail.Default.IsHourlyTrigger)
            //{
            //    WriteLog("Hourly Instock Service Started", " OnStart");
            //    _runFlag = true;
            //    _instockHourlyOutPutDataThread2 = new Thread(new ThreadStart(this.CollectHourlyInstockOutPutData2));
            //    _instockHourlyOutPutDataThread2.Start();
            //}

            //if (InstockAutoMail.Default.IsOBATrigger)
            //{
            //    WriteLog("OBA Service Started", " OnStart");
            //    _runFlag = true;
            //    _ObaOutputDataThread = new Thread(new ThreadStart(this.CollectOBAAutoOutPutData));
            //    _ObaOutputDataThread.Start();
            //}

            //if (InstockAutoMail.Default.IsVUTVAutoMail)
            //{
            //    WriteLog("VUTV Service Started", " OnStart");
            //    _runFlag = true;
            //    _VUTVShippedThread = new Thread(new ThreadStart(this.CollectVUTVShippmentData));
            //    _VUTVShippedThread.Start();

            //    _VUTVShippedAllThread = new Thread(new ThreadStart(this.CollectVUTVAllShippmentData));
            //    _VUTVShippedAllThread.Start();
            //}

            //if (InstockAutoMail.Default.IsMPKShippedTrigger)
            //{
            //    WriteLog("MPK Shipped Service Started", " OnStart");
            //    _runFlag = true;
            //    _MPKShippedThread = new Thread(new ThreadStart(this.CollectMPKShippedAutoOutPutData));
            //    _MPKShippedThread.Start();
            //}

            //if (InstockAutoMail.Default.IsPDLTrigger)
            //{
            //    WriteLog("PDL Service Started", " OnStart");
            //    _runFlag = true;
            //    _PdlOutputDataThread = new Thread(new ThreadStart(this.CollectPDLAutoOutPutData));
            //    _PdlOutputDataThread.Start();
            //}


            //if (InstockAutoMail.Default.IsSendSMS)
            //{
            //    WriteLog("SendPublicSMS Service Started", " OnStart");
            //    _runFlag = true;
            //    this._smsThread = new Thread(new ThreadStart(this.SendPublicSMS));
            //    this._smsThread.Start();
            //}

            //if (InstockAutoMail.Default.IsTimeOutTrigger)
            //{
            //    WriteLog("Timeout Service Started", " OnStart");
            //    _runFlag = true;
            //    _TimeOutDataThread2 = new Thread(new ThreadStart(this.Collect2HoursTimeGapData));
            //    _TimeOutDataThread2.Start();
            //    _TimeOutDataThread8 = new Thread(new ThreadStart(this.Collect8HoursTimeGapData));
            //    _TimeOutDataThread8.Start();
            //    _TimeOutDataThread24 = new Thread(new ThreadStart(this.Collect24HoursTimeGapData));
            //    _TimeOutDataThread24.Start();
            //}

        }

        private void CollectMpkInstockGapData()
        {
            WriteLog("Start.. CollectMpkInstockGapData Method " + DateTime.Now.ToString(), "CollectMpkInstockGapData");
            while (_runFlag)
            {
                int iHour = DateTime.Now.Hour;
                int iMinute = DateTime.Now.Minute;
                string timeNow2 = DateTime.Now.ToString("HH:mmtt");
                string mpkinstockNow = ConfigurationManager.AppSettings["MailTriggerMPKINSTOCK"].ToString();
                int iHours = int.Parse(InstockAutoMail.Default.IsHour);
                int iMinutes = int.Parse(InstockAutoMail.Default.IsMin);
                string subDate = DateTime.Now.ToString("yyyy MMMM dd HHmmtt");
                string eTime = iHour + ":" + iMinute;

                if (timeNow2 == mpkinstockNow || timeNow2 == "14:15PM" || timeNow2 == "22:15PM")
                {
                    WriteLog("Execution Time " + timeNow2, "ExecutionTime");
                    Mail ml = new Mail(Mail_User, Mail_Password, Mail_Host, Mail_Port);
                    DataTable dtMail = GetMailData("MPKINSTOCK").Tables[0];
                    WriteLog("Call Mail Id's " + DateTime.Now.ToString(), "CollectMpkInstockGapData");
                    DataRow row = dtMail.Rows[0];
                    AddMailAddress(ml, row);
                    ml.Subject = "(FUSE SYSTEM MAIL) MPK to INSTOCK GAP as on " + subDate;
                    ml.Body = "Hello Team<br/><br/> Here attached file as MPK to INSTOCK GAP Report " + subDate + ". <br/><br/><br/>With Best Regards!<br/>" + _CName + ".";

                    WriteLog(eTime, "CollectMpkInstockGapData");
                    try
                    {
                        DataTable finalData = new DataTable();

                        DataTable dtFamilies = GetOBAFamily("OQC").Tables[0];
                        if (dtFamilies.Rows.Count > 0)
                        {
                            for (int j = 0; j < dtFamilies.Rows.Count; j++)
                            {
                                string family = dtFamilies.Rows[j][0].ToString();
                                string schema = dtFamilies.Rows[j][1].ToString();
                                DataTable dtOQCOutPut = GetMpkInstockGapData(family, schema).Tables[0];

                                WriteLog("Get CollectMpkInstockGap Data " + DateTime.Now.ToString() + " " + family + " Count: " + dtOQCOutPut.Rows.Count, "CollectMpkInstockGapData");
                                if (dtOQCOutPut.Rows.Count > 0)
                                {
                                    finalData.Merge(dtOQCOutPut);

                                }
                                else
                                {
                                    continue;
                                }
                            }

                        }

                        if (finalData.Rows.Count > 0)
                        {
                            AddMPKINSTOCKGAPMailAttachment(ml, finalData, "MPKINSTOCKGAP", "Data");

                            ml.Send2();
                            WriteLog("Sent files to Mail Id's " + DateTime.Now.ToString(), "SendMail");
                        }

                    }
                    catch (Exception ex)
                    {
                        WriteLog("CollectMpkInstockGapData throw exception" + ex.Message, " CollectMpkInstockGapData");
                        _runFlag = false;
                    }
                }
                Thread.Sleep(60000);
            }
            WriteLog("End.. CollectMpkInstockGapData Method " + DateTime.Now.ToString(), "CollectMpkInstockGapData");
        }


        private void CollectOQCDisplayFailData()
        {
            WriteLog("Start.. CollectOQCDisplayFailData Method " + DateTime.Now.ToString(), "CollectOQCDisplayFailData");
            while (_runFlag)
            {
                int iHour = DateTime.Now.Hour;
                int iMinute = DateTime.Now.Minute;
                string timeNow2 = DateTime.Now.ToString("HH:mmtt");
                string OqctimeNow = ConfigurationManager.AppSettings["MailTriggerOQCDSPTime"].ToString();
                int iHours = int.Parse(InstockAutoMail.Default.IsHour);
                int iMinutes = int.Parse(InstockAutoMail.Default.IsMin);
                string subDate = DateTime.Now.AddDays(-1).ToString("yyyy MMMM dd");
                string eTime = iHour + ":" + iMinute;

                if (timeNow2 == OqctimeNow)
                {
                    Mail ml = new Mail(Mail_User, Mail_Password, Mail_Host, Mail_Port);
                    DataTable dtMail = GetMailData("OQCInfo").Tables[0];
                    WriteLog("Call Mail Id's " + DateTime.Now.ToString(), "CollectOQCDisplayFailData");
                    DataRow row = dtMail.Rows[0];
                    AddMailAddress(ml, row);
                    ml.Subject = "(FUSE SYSTEM MAIL) OQC Stage failed Mobiles Report " + subDate;
                    ml.Body = "Hello Team<br/><br/> Here attached file as OQC Stage failed Mobiles Report " + subDate + ". <br/><br/><br/>With Best Regards!<br/>" + _CName + ".";


                    WriteLog(eTime, "CollectOQCDisplayFailData");
                    try
                    {
                        DataTable finalData = new DataTable();

                        DataTable dtFamilies = GetOBAFamily("OQC").Tables[0];
                        if (dtFamilies.Rows.Count > 0)
                        {
                            for (int j = 0; j < dtFamilies.Rows.Count; j++)
                            {
                                string family = dtFamilies.Rows[j][0].ToString();
                                string schema = dtFamilies.Rows[j][1].ToString();
                                DataTable dtOQCOutPut = GetOQCDisplayFailData(family, schema).Tables[0];

                                WriteLog("Get OQC Data " + DateTime.Now.ToString() + " " + family + " Count: " + dtOQCOutPut.Rows.Count, "GetOBAOutPutData");
                                if (dtOQCOutPut.Rows.Count > 0)
                                {
                                    finalData.Merge(dtOQCOutPut);

                                }
                                else
                                {
                                    continue;
                                }
                            }

                        }

                        if (finalData.Rows.Count > 0)
                        {
                            AddOQCDISPAYFAILMailAttachment(ml, finalData, "OQCDISPLAYFAILData", "Data");

                            ml.Send2();
                            WriteLog("Sent files to Mail Id's " + DateTime.Now.ToString(), "SendMail");
                        }

                    }
                    catch (Exception ex)
                    {
                        WriteLog("CollectOQCDisplayFailData throw exception" + ex.Message, " CollectOQCDisplayFailData");
                        _runFlag = false;
                    }
                }
                Thread.Sleep(60000);
            }
            WriteLog("End.. CollectOQCDisplayFailData Method " + DateTime.Now.ToString(), "CollectOQCDisplayFailData");
        }
        private void CollectDashboardInstockFullData()
        {
            string timeNow = DateTime.Now.ToString("HHtt");
            try
            {
                if (timeNow == "08AM")
                {
                    WriteLog(timeNow, "Instock_Auto_Method_Executution_Started");
                    DataTable dtMFROutPut = con.getInstockFullData(DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd"));
                    //DataTable dtMFROutPut = GetInstockData().Tables[0]; //GetMFROutPutData().Tables[0];
                    if (dtMFROutPut.Rows.Count > 0)
                    {
                        DataTable dtMail = GetMailData("INSTOCKInfo").Tables[0];//MFROutPut
                        for (int i = 0; i < dtMail.Rows.Count; i++)
                        {
                            DataRow row = dtMail.Rows[i];
                            Mail ml = new Mail(Mail_User, Mail_Password, Mail_Host, Mail_Port);
                            AddMailAddress(ml, row);
                            ml.Subject = "(FUSE SYSTEM MAIL) " + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "Instock Output Data";
                            ml.Body = "Hi <br/><br/> Please check the attached <b>" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "</b> date Instock Report.<br/><br/><br/><br/>Regards<br/>BharathFIH-IT.";
                            if (AddMFRMailAttachment(ml, dtMFROutPut, "INSTOCKOutPutTemplate", "Data"))
                            {
                                ml.Send2();
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("CollectInstockAutoOutPutData throw exception" + ex.Message, " CollectInstockAutoOutPutData");
                instockAutoTimer.Stop();
                instockAutoTimer.Start();

            }

        }
        private void SendPublicSMS()
        {
            Thread.Sleep(15000);
            while (this._runFlag)
            {
                try
                {
                    DataSet unSendSMSRecord = this.GetUnSendSMSRecord();
                    bool flag = unSendSMSRecord.Tables.Count > 0;
                    if (flag)
                    {
                        WriteLog("Received Sending... Records: " + unSendSMSRecord.Tables.Count.ToString(), "SMS");
                        DataTable dataTable = unSendSMSRecord.Tables[0];
                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            DataRow dataRow = dataTable.Rows[i];
                            string strID = dataRow["ID"].ToString();
                            string text = dataRow["SUBJECT"].ToString();
                            byte[] bytes = (byte[])dataRow["SMS_BODY"];
                            string @string = Encoding.UTF8.GetString(bytes);
                            string[] array = dataRow["SEND_TO"].ToString().TrimEnd(new char[]
							{
								';'
							}).Split(new char[]
							{
								';'
							});
                            for (int j = 0; j < array.Length; j++)
                            {
                                this.sendsms(array[j], @string);
                            }
                            this.UpdateSMSActive(strID, "1");
                        }

                    }
                }
                catch (Exception ex)
                {
                    WriteLog("SendPublicSMS, error message: " + ex.Message, "SMS");
                }
                Thread.Sleep(60000);
            }
        }
        private void CollectVUTVShippmentData()
        {
            WriteLog("Start.. CollectVUTVShippmentData Method " + DateTime.Now.ToString(), "CollectVUTVShippmentData");
            while (_runFlag)
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1).Day;


                if (DateTime.Now.Day < lastDayOfMonth)
                {

                    int iHour = DateTime.Now.Hour;
                    int iMinute = DateTime.Now.Minute;
                    string timeNow2 = DateTime.Now.ToString("HH:mmtt");
                    string ObatimeNow = ConfigurationManager.AppSettings["MailTriggeVUTVTime"].ToString();

                    DateTime durtime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                    string fulldurtime = durtime.ToString("dd MMMM yyyy") + " 06:00" + " to " + DateTime.Now.ToString("dd MMMM yyyy") + " 06:00";
                    int iHours = int.Parse(InstockAutoMail.Default.IsHour);
                    int iMinutes = int.Parse(InstockAutoMail.Default.IsMin);
                    string eTime = iHour + ":" + iMinute;
                    Mail ml = new Mail(Mail_User, Mail_Password, Mail_Host, Mail_Port);
                    DataTable dtMail = GetMailData("VUTVShipped").Tables[0];
                    WriteLog("Call Mail Id's " + DateTime.Now.ToString(), "CollectVUTVShippmentData");
                    DataRow row = dtMail.Rows[0];
                    AddMailAddress(ml, row);
                    ml.Subject = "(FUSE SYSTEM MAIL) Cumulative Daily Month Wise Production Data";
                    //ml.Body = "Hello Team<br/><br/> Here attached Cumulative Daily Production Data Upto 06:00AM Data <br/><br/><br/><br/>Regards<br/>BharathFIH-IT.";
                    ml.Body = "Hello Team<br/><br/> Here attached Cumulative Daily Production Data From " + fulldurtime + ". <br/><br/><br/><br/>Regards<br/>" + _CName + ".";
                    if (timeNow2 == ObatimeNow)
                    {
                        WriteLog(eTime, "CollectVUTVShippmentData");
                        try
                        {
                            DataTable dtFamilies = GetTVFamily().Tables[0];

                            if (dtFamilies.Rows.Count > 0)
                            {
                                //for (int j = 0; j < dtFamilies.Rows.Count; j++)
                                //{
                                string family = dtFamilies.Rows[0][0].ToString();
                                string schema = dtFamilies.Rows[0][1].ToString();
                                DataTable dtTVOutPut = GetVUTVShippmenttData(family, schema).Tables[0];

                                //var query = (from t in dtTVOutPut.AsEnumerable()
                                //             group t by new { ModelCode = t.Field<string>("MODEL_CODE") }
                                //                 into grp
                                //                 select new
                                //                 {
                                //                     grp.Key.ModelCode

                                //                 }).ToList();


                                //foreach (var item in query)
                                //{
                                //    DataTable dtTVOutPut2 = dtTVOutPut.AsEnumerable().Where(a => a.Field<string>("MODEL_CODE") == item.ModelCode).CopyToDataTable();

                                //}

                                WriteLog("Get VUTV Data " + DateTime.Now.ToString() + " " + family + " Count: " + dtTVOutPut.Rows.Count, "CollectVUTVShippmentData");
                                if (dtTVOutPut.Rows.Count > 0)
                                {
                                    AddVUTVMailAttachment(ml, dtTVOutPut, "VUTVShippmentTemplate", "Data");
                                }
                                else
                                {
                                    ml.Body += "<br/><br/><b> Note: Today shippment process not done. </b>";
                                }
                                //}

                            }

                            ml.Send2();
                            WriteLog("Sent files to Mail Id's " + DateTime.Now.ToString(), "SendMail");

                        }
                        catch (Exception ex)
                        {
                            WriteLog("CollectVUTVShippmentData throw exception" + ex.Message, " CollectVUTVShippmentData");
                            _runFlag = false;
                        }
                    }
                    Thread.Sleep(120000);
                }
            }
            WriteLog("End.. CollectVUTVShippmentData Method " + DateTime.Now.ToString(), "CollectVUTVShippmentData");
        }
        private void CollectVUTVAllShippmentData()
        {
            WriteLog("Start.. CollectVUTVAllShippmentData Method " + DateTime.Now.ToString(), "CollectVUTVAllShippmentData");
            while (_runFlag)
            {
                var firstDayOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1).Day;


                if (DateTime.Now.Day == lastDayOfMonth)
                {

                    int iHour = DateTime.Now.Hour;
                    int iMinute = DateTime.Now.Minute;
                    string timeNow2 = DateTime.Now.ToString("HH:mmtt");
                    string ObatimeNow = ConfigurationManager.AppSettings["MailTriggeVUTVTime"].ToString();

                    DateTime durtime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                    string fulldurtime = durtime.ToString("dd MMMM yyyy") + " 06:00" + " to " + DateTime.Now.ToString("dd MMMM yyyy") + " 06:00";
                    int iHours = int.Parse(InstockAutoMail.Default.IsHour);
                    int iMinutes = int.Parse(InstockAutoMail.Default.IsMin);
                    string eTime = iHour + ":" + iMinute;
                    Mail ml = new Mail(Mail_User, Mail_Password, Mail_Host, Mail_Port);
                    DataTable dtMail = GetMailData("VUTVShipped").Tables[0];
                    WriteLog("Call Mail Id's " + DateTime.Now.ToString(), "CollectVUTVAllShippmentData");
                    DataRow row = dtMail.Rows[0];
                    AddMailAddress(ml, row);
                    ml.Subject = "(FUSE SYSTEM MAIL) Cumulative Overall Production Data";
                    //ml.Body = "Hello Team<br/><br/> Here attached Cumulative Daily Production Data Upto 06:00AM Data <br/><br/><br/><br/>Regards<br/>BharathFIH-IT.";
                    ml.Body = "Hello Team<br/><br/> Here attached Cumulative Overall Production Data. <br/><br/><br/><br/>Regards<br/>" + _CName + ".";
                    if (timeNow2 == ObatimeNow)
                    {
                        if (DateTime.Now.Day == lastDayOfMonth)
                        {

                            WriteLog(eTime, "CollectVUTVAllShippmentData");
                            try
                            {
                                DataTable dtFamilies = GetTVFamily().Tables[0];

                                if (dtFamilies.Rows.Count > 0)
                                {
                                    //for (int j = 0; j < dtFamilies.Rows.Count; j++)
                                    //{
                                    string family = dtFamilies.Rows[0][0].ToString();
                                    string schema = dtFamilies.Rows[0][1].ToString();
                                    DataTable dtTVOutPut = GetVUTVAllShippmenttData(family, schema).Tables[0];

                                    //var query = (from t in dtTVOutPut.AsEnumerable()
                                    //             group t by new { ModelCode = t.Field<string>("MODEL_CODE") }
                                    //                 into grp
                                    //                 select new
                                    //                 {
                                    //                     grp.Key.ModelCode

                                    //                 }).ToList();


                                    //foreach (var item in query)
                                    //{
                                    //    DataTable dtTVOutPut2 = dtTVOutPut.AsEnumerable().Where(a => a.Field<string>("MODEL_CODE") == item.ModelCode).CopyToDataTable();

                                    //}

                                    WriteLog("Get VUTV Data " + DateTime.Now.ToString() + " " + family + " Count: " + dtTVOutPut.Rows.Count, "CollectVUTVAllShippmentData");
                                    if (dtTVOutPut.Rows.Count > 0)
                                    {
                                        AddVUTVMailAttachment(ml, dtTVOutPut, "VUTVShippmentTemplate", "Data");
                                    }
                                    else
                                    {
                                        ml.Body += "<br/><br/><b> Note: Today shippment process not done. </b>";
                                    }
                                    //}

                                }

                                ml.Send2();
                                WriteLog("Sent files to Mail Id's " + DateTime.Now.ToString(), "SendMail");

                            }
                            catch (Exception ex)
                            {
                                WriteLog("CollectVUTVAllShippmentData throw exception" + ex.Message, " CollectVUTVAllShippmentData");
                                _runFlag = false;
                            }
                        }
                        Thread.Sleep(60000);
                    }
                }
            }
            WriteLog("End.. CollectVUTVAllShippmentData Method " + DateTime.Now.ToString(), "CollectVUTVAllShippmentData");
        }
        private void CollectOBAAutoOutPutData()
        {
            WriteLog("Start.. CollectOBAAutoOutPutData Method " + DateTime.Now.ToString(), "CollectOBAAutoOutPutData");
            while (_runFlag)
            {
                int iHour = DateTime.Now.Hour;
                int iMinute = DateTime.Now.Minute;
                string timeNow2 = DateTime.Now.ToString("HH:mmtt");
                string ObatimeNow = ConfigurationManager.AppSettings["MailTriggerOBATime"].ToString();
                string ObatimeShiftA = ConfigurationManager.AppSettings["MailTriggerOBATimeShiftA"].ToString();
                string ObatimeShiftB = ConfigurationManager.AppSettings["MailTriggerOBATimeShiftB"].ToString();
                int iHours = int.Parse(InstockAutoMail.Default.IsHour);
                int iMinutes = int.Parse(InstockAutoMail.Default.IsMin);
                string eTime = iHour + ":" + iMinute;
                Mail ml = new Mail(Mail_User, Mail_Password, Mail_Host, Mail_Port);
                DataTable dtMail = GetMailData("OBAInfo").Tables[0];
                WriteLog("Call Mail Id's " + DateTime.Now.ToString(), "CollectOBAAutoOutPutData");
                DataRow row = dtMail.Rows[0];
                AddMailAddress(ml, row);
                ml.Subject = "(FUSE SYSTEM MAIL) OBA mobiles reversed to OQC stage. Not scanned in OQC more than 24 hours";
                ml.Body = "Hello Team<br/><br/> Here attached file as OBA mobiles reversed to OQC stage. Not scanned in OQC more than 24 hours....! please pass the next station. <br/><br/><br/><br/>Regards<br/>" + _CName + ".";

                if ((timeNow2 == ObatimeNow) || (timeNow2 == ObatimeShiftA) || (timeNow2 == ObatimeShiftB))//shift wise
                {
                    WriteLog(eTime, "CollectOBAAutoOutPutData");
                    try
                    {
                        DataTable dtFamilies = GetOBAFamily("OBA").Tables[0];
                        if (dtFamilies.Rows.Count > 0)
                        {
                            for (int j = 0; j < dtFamilies.Rows.Count; j++)
                            {
                                string family = dtFamilies.Rows[j][0].ToString();
                                string schema = dtFamilies.Rows[j][1].ToString();
                                DataTable dtMFROutPut = GetOBAOutPutData(family, schema).Tables[0];

                                WriteLog("Get OBA Data " + DateTime.Now.ToString() + " " + family + " Count: " + dtMFROutPut.Rows.Count, "GetOBAOutPutData");
                                if (dtMFROutPut.Rows.Count > 0)
                                {
                                    AddOBAMailAttachment(ml, dtMFROutPut, "OBAOutPutTemplate", "Data");
                                }
                                else
                                {
                                    continue;
                                }
                            }
                        }
                        ml.Send2();
                        WriteLog("Sent files to Mail Id's " + DateTime.Now.ToString(), "SendMail");
                    }
                    catch (Exception ex)
                    {
                        WriteLog("CollectOBAAutoOutPutData throw exception" + ex.Message, " CollectOBAAutoOutPutData");
                        _runFlag = false;
                    }
                }
                Thread.Sleep(60000);
            }
            WriteLog("End.. CollectOBAAutoOutPutData Method " + DateTime.Now.ToString(), "CollectOBAAutoOutPutData");
        }

        private void CollectMPKShippedAutoOutPutData()
        {
            WriteLog("Start.. CollectMPKShippedAutoOutPutData Method " + DateTime.Now.ToString(), "CollectMPKShippedAutoOutPutData");
            while (_runFlag)
            {
                int iHour = DateTime.Now.Hour;
                int iMinute = DateTime.Now.Minute;
                string timeNow2 = DateTime.Now.ToString("HH:mmtt");
                string mpktimeNow = ConfigurationManager.AppSettings["MailTriggerMPKTime"].ToString();
                int iHours = int.Parse(InstockAutoMail.Default.IsHour);
                int iMinutes = int.Parse(InstockAutoMail.Default.IsMin);
                string eTime = iHour + ":" + iMinute;
                Mail ml = new Mail(Mail_User, Mail_Password, Mail_Host, Mail_Port);
                DataTable dtMail = GetMailData("MPKShipped").Tables[0];
                WriteLog("Call Mail Id's " + DateTime.Now.ToString(), "CollectMPKShippedAutoOutPutData");
                DataRow row = dtMail.Rows[0];
                AddMailAddress(ml, row);
                ml.Subject = "(FUSE SYSTEM MAIL) MPK Boxes Finished, But Not shipped with in 72 hours";
                ml.Body = "Hello Team<br/><br/> Here attached file as MPK Boxes Finished, But Not shipped with in 72 hours....! please check. <br/><br/><br/><br/>Regards<br/>BharathFIH-IT.";
                if (timeNow2 == mpktimeNow)
                {
                    WriteLog(eTime, "CollectMPKShippedAutoOutPutData");
                    try
                    {
                        DataTable dtFamilies = GetMPKFamily().Tables[0];
                        if (dtFamilies.Rows.Count > 0)
                        {
                            for (int j = 0; j < dtFamilies.Rows.Count; j++)
                            {
                                string family = dtFamilies.Rows[j]["CODE_INFO"].ToString();
                                string schema = dtFamilies.Rows[j]["CODE_INFO_NAME"].ToString();
                                DataTable dtMFROutPut = GetMPKNotShippedData(family, schema).Tables[0];
                                WriteLog("Get MPK Finished Data " + DateTime.Now.ToString() + " " + family + " Count: " + dtMFROutPut.Rows.Count, "CollectMPKShippedAutoOutPutData");
                                if (dtMFROutPut.Rows.Count > 0)
                                {
                                    if (!AddMPKShippedMailAttachment(ml, dtMFROutPut, "MPKFinishedData", "Data"))
                                        continue;
                                }



                                else
                                {
                                    continue;
                                }
                            }

                        }

                        ml.Send2();
                        WriteLog("Sent files to Mail Id's " + DateTime.Now.ToString(), "SendMail");

                    }
                    catch (Exception ex)
                    {
                        WriteLog("CollectMPKShippedAutoOutPutData throw exception" + ex.Message, " CollectMPKShippedAutoOutPutData");
                        _runFlag = false;
                    }
                }
                Thread.Sleep(60000);
            }
            WriteLog("End.. CollectMPKShippedAutoOutPutData Method " + DateTime.Now.ToString(), "CollectMPKShippedAutoOutPutData");
        }
        private void CollectPDLAutoOutPutData()
        {
            WriteLog("Start.. CollectPDLAutoOutPutData Method " + DateTime.Now.ToString(), "CollectPDLAutoOutPutData");
            while (_runFlag)
            {
                int iHour = DateTime.Now.Hour;
                int iMinute = DateTime.Now.Minute;
                string timeNow2 = DateTime.Now.ToString("HH:mmtt");
                string ObatimeNow = ConfigurationManager.AppSettings["MailTriggerPDLTime"].ToString();
                int iHours = int.Parse(InstockAutoMail.Default.IsHour);
                int iMinutes = int.Parse(InstockAutoMail.Default.IsMin);
                string eTime = iHour + ":" + iMinute;
                Mail ml = new Mail(Mail_User, Mail_Password, Mail_Host, Mail_Port);
                DataTable dtMail = GetMailDataPdl("PDLInfo").Tables[0];
                WriteLog("Call Mail Id's " + DateTime.Now.ToString(), "CollectPDLAutoOutPutData");
                DataRow row = dtMail.Rows[0];
                AddMailAddress(ml, row);
                ml.Subject = "(FUSE SYSTEM MAIL) PDL Passed mobiles, Not scanned in OQC more than 24 hours";
                ml.Body = "Hello Team<br/><br/> Here attached file as PDL Passed mobiles, Not scanned in OQC more than 24 hours....! please pass the next station. <br/><br/><br/><br/>Regards<br/>" + _CName + ".";
                if (timeNow2 == ObatimeNow)
                {
                    WriteLog(ObatimeNow, "CollectPDLAutoOutPutData");
                    try
                    {
                        DataTable dtFamilies = GetOBAFamily("PDL").Tables[0];
                        if (dtFamilies.Rows.Count > 0)
                        {
                            for (int j = 0; j < dtFamilies.Rows.Count; j++)
                            {
                                string family = dtFamilies.Rows[j][0].ToString();
                                string schema = dtFamilies.Rows[j][1].ToString();
                                DataTable dtMFROutPut = GetPDLOutPutData(family, schema).Tables[0];

                                WriteLog("Get PDL Data " + DateTime.Now.ToString() + " " + family + " Count: " + dtMFROutPut.Rows.Count, "CollectPDLAutoOutPutData");
                                if (dtMFROutPut.Rows.Count > 0)
                                {
                                    AddOBAMailAttachment2(ml, dtMFROutPut, "PDLOutPutTemplate", "Data");
                                }
                                else
                                {
                                    continue;
                                }
                            }

                        }

                        ml.Send2();
                        WriteLog("Sent files to Mail Id's " + DateTime.Now.ToString(), "SendMail");

                    }
                    catch (Exception ex)
                    {
                        WriteLog("CollectPDLAutoOutPutData throw exception" + ex.Message, " CollectPDLAutoOutPutData");
                        _runFlag = false;
                    }
                }
                Thread.Sleep(60000);
            }
            WriteLog("End.. CollectPDLAutoOutPutData Method " + DateTime.Now.ToString(), "CollectPDLAutoOutPutData");
        }
        private DataSet GetOBAFamily(string type)
        {
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                string strSql = "SELECT CODE_INFO, CODE_INFO_NAME FROM MASTER_USER.CODE_INFO Where CODE_ID IN('2021') AND CODE_INFO_NAME_DESC='" + type + "' ORDER BY CODE_INFO";
                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetOBAFamily, error message: " + ex.Message, " OBAFAMILY");
                }
            }
            return myds;
        }
        private DataSet GetTVFamily()
        {
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                string strSql = "SELECT CODE_INFO, CODE_INFO_NAME FROM MASTER_USER.CODE_INFO Where CODE_ID IN('2021') AND CODE_INFO_NAME_DESC='VUTV'";
                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetOBAFamily, error message: " + ex.Message, " OBAFAMILY");
                }
            }
            return myds;
        }
        private DataSet GetUnSendSMSRecord()
        {
            string strSbject = System.Configuration.ConfigurationManager.AppSettings["MessageSubject"].ToString();
            string strQuery = string.Empty;
            string[] arr = strSbject.Split(',');
            foreach (var item in arr)
            {
                strQuery += "'" + item + "',";
            }
            //strQuery = "'"+strQuery.TrimEnd(',');
            strQuery = strQuery.Remove(strQuery.Length - 1, 1);
            DataSet dataSet = new DataSet();
            using (OracleConnection oracleConnection = new OracleConnection(this._DBConnStr))
            {
                string cmdText = " SELECT * FROM LINE_MONITOR.AUTO_SMS A WHERE A.ACTIVE IN ( 0 ) and A.SUBJECT IN(" + strQuery + ") ";
                OracleCommand selectCommand = new OracleCommand(cmdText, oracleConnection);
                try
                {
                    oracleConnection.Open();
                    OracleDataAdapter oracleDataAdapter = new OracleDataAdapter(selectCommand);
                    oracleDataAdapter.Fill(dataSet);
                }
                catch (Exception ex)
                {
                    WriteLog("GetUnSendSMSRecord, error message: " + ex.Message, "SMS");
                }
            }
            return dataSet;
        }

        private void UpdateSMSActive(string strID, string strActive)
        {
            string cmdText = string.Concat(new string[]
			{
				" UPDATE LINE_MONITOR.AUTO_SMS SET ACTIVE = '",
				strActive,
				"' WHERE ID='",
				strID,
				"' "
			});
            using (OracleConnection oracleConnection = new OracleConnection(this._DBConnStr))
            {
                oracleConnection.Open();
                OracleCommand oracleCommand = new OracleCommand(cmdText, oracleConnection);
                oracleCommand.ExecuteNonQuery();
                oracleConnection.Close();
            }
        }
        private DataSet GetMPKFamily()
        {
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                //string strSql = "SELECT FAMILY_ID, SCHEMA_NAME,SERVER_DESC,DB_NAME FROM SFIS1.V_FAMILY_DB_CONFIG_ALL Where B2B_FLAG=1 AND DB_NAME='INDB01' ";
                string strSql = "SELECT CODE_INFO, CODE_INFO_NAME FROM MASTER_USER.CODE_INFO Where CODE_ID IN('2021') AND CODE_INFO_NAME_DESC=3";
                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetOBAFamily, error message: " + ex.Message, " OBAFAMILY");
                }
            }
            return myds;
        }
        private DataSet GetOBAFamily2(int sequence)
        {
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                string strSql = "SELECT CODE_INFO, CODE_INFO_NAME FROM MASTER_USER.CODE_INFO Where CODE_ID IN('2021') AND CODE_INFO_NAME_DESC=" + sequence;
                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetOBAFamily, error message: " + ex.Message, " OBAFAMILY");
                }
            }
            return myds;
        }
        public DataSet GetVUTVShippmenttDataProc()
        {
            DataSet myDs = new DataSet();
            using (OracleConnection orcn = new OracleConnection(_DBConnStr))
            {
                OracleCommand cmd = new OracleCommand("yield_report.get_mfr_output_data", orcn);
                OracleDataAdapter oda = new OracleDataAdapter(cmd);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("p_time_start", OracleDbType.Varchar2, 50).Value = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 08:00");
                cmd.Parameters.Add("p_time_end", OracleDbType.Varchar2, 50).Value = DateTime.Now.ToString("yyyy-MM-dd 08:00");
                cmd.Parameters.Add("o_item_data", OracleDbType.RefCursor).Direction = ParameterDirection.Output;
                try
                {
                    orcn.Open();
                    oda.Fill(myDs);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMFROutPutData, error message: " + ex.Message, "GetMFROutPutData");
                }
            }
            return myDs;
        }
        private DataSet GetVUTVShippmenttData(string family, string schema)
        {
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                //string strSql = "WITH a " +
                //                     "   AS (  SELECT rh.model_code, " +
                //                      "               ac.psn AS tv_sn, " +
                //                       "               MAX (DECODE (ac.name, 'LCM_SN', ac.VALUE, '')) lcm_sn, " +
                //                         "             MAX (DECODE (ac.name, 'MB_SN', ac.VALUE, '')) mb_sn, " +
                //                          "            pe.create_time fisrtscan_time " +
                //                           "      FROM " + schema + ".acc_code ac " +
                //                            "          JOIN " + schema + ".routing_history rh ON rh.psn = ac.psn " +
                //               "                      JOIN " + schema + ".pack_history ph " +
                //               "                         ON     ph.psn = ac.psn " +
                //               "                            AND ph.active = 1 " +
                //               "                            AND ph.pack_type = 'ATO' " +
                //               "                      JOIN " + schema + ".pack_entry pe ON pe.pack_id = ph.masterpack_number " +
                //               "                WHERE pe.create_time <= TRUNC (SYSDATE + 1) + 6 / 24 " +
                //               "             GROUP BY ac.psn, rh.model_code, pe.create_time) " +
                //               "      SELECT a.model_code, " +
                //               "             TO_CHAR (a.fisrtscan_time, 'YYYY-MM-DD') AS fisrtscan_time, " +
                //               "             a.tv_sn, " +
                //               "             a.lcm_sn, " +
                //               "             a.mb_sn, " +
                //               "             ac2.VALUE AS oc_sn " +
                //               "        FROM a JOIN " + schema + ".acc_code ac2 ON a.lcm_sn = ac2.psn AND ac2.name = 'OC_SN' " +
                //               "    ORDER BY a.fisrtscan_time";


                string strSql = "SELECT pe.family_id,pe.model_code, " +
                                         "  rh.create_time AS fisrtscan_time, " +
                                         "  ph.psn as tv_sn,                                                                 " +
                                         "  DECODE (ac.name, 'LCM_SN', ac.VALUE) AS lcm_sn,                         " +
                                         "  ac2.VALUE AS mb_sn,                                                     " +
                                         "  ac4.VALUE AS sb_sn,                                                     " +
                                         "  ac5.VALUE AS oc_sn                                                      " +
                                     " FROM " + schema + ".pack_entry pe                                                      " +
                                     "      LEFT JOIN " + schema + ".pack_history ph                                          " +
                                     "         ON     pe.pack_id = ph.masterpack_number                             " +
                                     "            AND ph.pack_type = 'ATO'                                          " +
                                     "            AND ph.active = 1                                                 " +
                                     "            AND ph.masterpack_number IS NOT NULL                              " +
                                     "      LEFT JOIN (SELECT PSN, NAME, VALUE                                      " +
                                     "                   FROM " + schema + ".ACC_CODE                                              " +
                                     "                  WHERE NAME = 'LCM_SN'                                       " +
                                     "                 UNION                                                        " +
                                     "                 SELECT PSN, DECODE (NAME, 'LCM', 'LCM_SN') AS NAME, VALUE    " +
                                     "                   FROM " + schema + ".ACC_CODE                                              " +
                                     "                  WHERE NAME = 'LCM') ac                                      " +
                                     "         ON ph.psn = ac.psn AND ac.name IN ('LCM_SN')                         " +
                                     "      LEFT JOIN " + schema + ".acc_code ac2 ON ph.psn = ac2.psn AND ac2.name = 'MB_SN'  " +
                                     "      LEFT JOIN " + schema + ".acc_code ac4 ON ph.psn = ac4.psn AND ac4.name = 'SB_SN'  " +
                                     "      LEFT JOIN " + schema + ".acc_code ac5                                             " +
                                     "         ON ac.VALUE = ac5.psn AND ac5.name = 'OC_SN'                         " +
                                     "      LEFT JOIN                                                               " +
                                     "      (SELECT psn,                                                            " +
                                     "              station_group,                                                  " +
                                     "              create_time,                                                    " +
                                     "              status,                                                         " +
                                     "              ROW_NUMBER ()                                                   " +
                                     "                 OVER (PARTITION BY psn ORDER BY create_time DESC)            " +
                                     "                 AS rown                                                      " +
                                     "         FROM " + schema + ".routing_history                                            " +
                                     "        WHERE station_group = 'FAI' AND status = 'P') rh                      " +
                                     "         ON ph.psn = rh.psn                                                   " +
                                     " WHERE pe.active IN (1, 2) and PE.CREATE_TIME BETWEEN TO_DATE (TO_CHAR (TRUNC (SYSDATE, 'MM'), " +
                                     "            'YYYY-MM-DD') || ' 06:00:00','YYYY-MM-DD HH24:MI;SS') AND TO_DATE (TO_CHAR (TRUNC (SYSDATE), " +
                                     "            'YYYY-MM-DD')|| ' 06:00:00','YYYY-MM-DD HH24:MI;SS') AND rh.rown = 1";

                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }
        private DataSet GetVUTVAllShippmenttData(string family, string schema)
        {
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {

                string strSql = "SELECT pe.family_id,pe.model_code, " +
                                         "  rh.create_time AS fisrtscan_time, " +
                                         "  ph.psn as tv_sn,                                                                 " +
                                         "  DECODE (ac.name, 'LCM_SN', ac.VALUE) AS lcm_sn,                         " +
                                         "  ac2.VALUE AS mb_sn,                                                     " +
                                         "  ac4.VALUE AS sb_sn,                                                     " +
                                         "  ac5.VALUE AS oc_sn                                                      " +
                                     " FROM " + schema + ".pack_entry pe                                                      " +
                                     "      LEFT JOIN " + schema + ".pack_history ph                                          " +
                                     "         ON     pe.pack_id = ph.masterpack_number                             " +
                                     "            AND ph.pack_type = 'ATO'                                          " +
                                     "            AND ph.active = 1                                                 " +
                                     "            AND ph.masterpack_number IS NOT NULL                              " +
                                     "      LEFT JOIN (SELECT PSN, NAME, VALUE                                      " +
                                     "                   FROM " + schema + ".ACC_CODE                                              " +
                                     "                  WHERE NAME = 'LCM_SN'                                       " +
                                     "                 UNION                                                        " +
                                     "                 SELECT PSN, DECODE (NAME, 'LCM', 'LCM_SN') AS NAME, VALUE    " +
                                     "                   FROM " + schema + ".ACC_CODE                                              " +
                                     "                  WHERE NAME = 'LCM') ac                                      " +
                                     "         ON ph.psn = ac.psn AND ac.name IN ('LCM_SN')                         " +
                                     "      LEFT JOIN " + schema + ".acc_code ac2 ON ph.psn = ac2.psn AND ac2.name = 'MB_SN'  " +
                                     "      LEFT JOIN " + schema + ".acc_code ac4 ON ph.psn = ac4.psn AND ac4.name = 'SB_SN'  " +
                                     "      LEFT JOIN " + schema + ".acc_code ac5                                             " +
                                     "         ON ac.VALUE = ac5.psn AND ac5.name = 'OC_SN'                         " +
                                     "      LEFT JOIN                                                               " +
                                     "      (SELECT psn,                                                            " +
                                     "              station_group,                                                  " +
                                     "              create_time,                                                    " +
                                     "              status,                                                         " +
                                     "              ROW_NUMBER ()                                                   " +
                                     "                 OVER (PARTITION BY psn ORDER BY create_time DESC)            " +
                                     "                 AS rown                                                      " +
                                     "         FROM " + schema + ".routing_history                                            " +
                                     "        WHERE station_group = 'FAI' AND status = 'P') rh                      " +
                                     "         ON ph.psn = rh.psn                                                   " +
                                     " WHERE pe.active IN (1, 2)  AND rh.rown = 1";

                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }

        private DataSet GetMpkInstockGapData(string family, string schema)
        {
            DataSet myds = new DataSet();

            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                string strSql = "  WITH aa"
                                   + "     AS (SELECT *                                                                "
                                   + "                FROM (SELECT ph.*, pe.pack_id                                    "
                                   + "                        FROM " + schema + ".pack_history ph                                     "
                                   + "                             LEFT JOIN " + schema + ".PACK_ENTRY pe                             "
                                   + "                                ON     PH.MASTERPACK_NUMBER = PE.PACK_ID         "
                                   + "                                   AND PE.ACTIVE > 0                             "
                                   + "                       WHERE     ph.pallet_number IS NULL                        "
                                   + "                             AND ph.pack_type = 'ATO'                            "
                                   + "                             AND ph.active = 1                                   "
                                   + "                             AND masterpack_number IS NOT NULL                   "
                                   + "                             AND (SYSDATE - ph.pack_time) * 24 >= 4)            "
                                   + "               WHERE pack_id IS NULL)                                            "
                                   + "     SELECT wo.family_id,                                                        "
                                   + "            wo.model_code,                                                       "
                                   + "            aa.psn,                                                              "
                                   + "            sc.VALUE imei,                                                       "
                                   + "            sc1.VALUE imei2,                                                     "
                                   + "            aa.pack_time,                                                        "
                                   + "            ROUND ( (SYSDATE - aa.pack_time) * 24, 2) total_hrs,                  "
                                   + "           CEIL (ROUND ( (SYSDATE - aa.pack_time) * 24, 2) / 24) total_days      "
                                   + "       FROM aa                                                                   "
                                   + "            LEFT JOIN " + schema + ".ssn_code sc ON aa.psn = sc.psn AND sc.name = 'IMEI'        "
                                   + "            LEFT JOIN " + schema + ".ssn_code sc1 ON aa.psn = sc1.psn AND sc1.name = 'IMEI2'    "
                                   + "            LEFT JOIN work_order wo ON aa.order_number = wo.order_number         ";

                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }
        private DataSet GetOQCDisplayFailData(string family, string schema)
        {
            DataSet myds = new DataSet();
            string p_time_start = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 06:00");
            string p_time_end = DateTime.Now.ToString("yyyy-MM-dd 06:00");

            string strfailItems = System.Configuration.ConfigurationManager.AppSettings["MailTriggerOQCDSPFITEMS"].ToString();

            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                //string strSql = "SELECT '" + family + "' AS FAMILY_ID,                                           " +
                //                "                      SN AS PSN,                                              " +
                //                "                      CREATE_TIME AS FAIL_TIME,                               " +
                //                "                      SYMPTON_FAILTURE_GROUP AS FAILTYPE_GROUP,               " +
                //                "                      SYMPTON_FAILTURE_ITEM AS FAILTYPE_GROUP_ITEM            " +
                //                "                 FROM " + schema + ".HISTORY_ITEM                           " +
                //                "                WHERE     OCCURRED_STATION_GROUP = 'OQC'                      "+
                //                "                  AND CREATE_TIME BETWEEN TO_DATE ('" + p_time_start + "', 'YYYY-MM-DD HH24:MI:SS') " +
                //                "                  AND TO_DATE ('" + p_time_end + "', 'YYYY-MM-DD HH24:MI:SS')";

                //string strSql = "SELECT DISTINCT '" + family + "' AS FAMILY_ID,rh.psn, AFM.FUNCTION_NAME, REPLACE (TRIM (AFM.TEST_NAME), '  ', ' ') as TEST_NAME FROM " + schema + ".routing_history rh JOIN " + schema + ".all_failure_measuerment afm ON RH.RESULT_ID = AFM.RESULT_ID WHERE rh.create_time BETWEEN TO_DATE ('" + p_time_start + "','YYYY-MM-DD HH24:MI:SS') AND TO_DATE ('" + p_time_end + "','YYYY-MM-DD HH24:MI:SS') AND rh.station_group = 'OQC' AND rh.status = 'F' AND rh.times_flag = 1 and afm.test_name in('Display Black Dot','Display Black  Patch','Display Black Patch','Display Black Dot ','Display White Patch','Display White Dot','Display White Dot ','Display White  Patch')";
                string strSql = "SELECT DISTINCT '" + family + "' AS FAMILY_ID,rh.psn, AFM.FUNCTION_NAME, REPLACE (TRIM (AFM.TEST_NAME), '  ', ' ') as TEST_NAME FROM " + schema + ".routing_history rh JOIN " + schema + ".all_failure_measuerment afm ON RH.RESULT_ID = AFM.RESULT_ID WHERE rh.create_time BETWEEN TO_DATE ('" + p_time_start + "','YYYY-MM-DD HH24:MI:SS') AND TO_DATE ('" + p_time_end + "','YYYY-MM-DD HH24:MI:SS') AND rh.station_group = 'OQC' AND rh.status = 'F' AND rh.times_flag = 1 and afm.test_name in(" + strfailItems + ")";

                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }

        private DataSet GetOBAOutPutData(string family, string schema)
        {
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                string strSql = "WITH aa"
                                    + " AS (SELECT *"
                                    + "FROM " + schema + ".ROUTING_WIP "
                                    + "WHERE psn IN (SELECT DISTINCT psn "
                                    + "FROM " + schema + ".PACK_HISTORY "
                                         + " WHERE     masterpack_number IN (SELECT SN "
                                                           + " FROM " + schema + ".REVERSE_HISTORY "
                                                         + " WHERE REVERSE_TYPE IN ('Customer OBA','Internal OBA')) "
                                                        + " AND active = 2)) "
                                    + " SELECT '" + family + "' as family_id" + " ,aa.psn,sc.VALUE AS imei,aa.station_group AS current_station_group, "
                                           + " aa.last_modified_time AS obadatetime "
                                    + " FROM aa LEFT JOIN " + schema + ".ssn_code sc ON aa.psn = sc.psn AND sc.name = 'IMEI' "
                                           + " WHERE station_group = 'OQC' AND (SYSDATE - LAST_MODIFIED_TIME) * 24 >= 24 ";
                //string strSql = "WITH aa"
                //                    +" AS (SELECT *"
                //                    +"FROM K7A.ROUTING_WIP "
                //                    +"WHERE psn IN (SELECT DISTINCT psn "
                //                    +"FROM K7A.PACK_HISTORY "
                //                         +" WHERE     masterpack_number IN (SELECT SN "
                //                                           +" FROM K7A.REVERSE_HISTORY "
                //                                         +" WHERE REVERSE_TYPE IN ('Customer OBA','Internal OBA')) "
                //                                        +" AND active = 2)) "
                //                    +" SELECT "+"'K7A'"+" as family_id"+" ,aa.psn,sc.VALUE AS imei,aa.station_group AS current_station_group, "
                //                           +" aa.last_modified_time AS obadatetime "
                //                    +" FROM aa LEFT JOIN K7A.ssn_code sc ON aa.psn = sc.psn AND sc.name = 'IMEI' "
                //                           + " WHERE station_group = 'OQC' AND (SYSDATE - LAST_MODIFIED_TIME) * 24 >= 24 ";

                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }
        private DataSet GetMPKNotShippedData(string family, string schema)
        {
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {



                //string strSql = "WITH AA "
                //                   + "  AS (SELECT * "
                //                     + "      FROM (SELECT DISTINCT pe.pack_id, "
                //                     + "                            pe.model_code, "
                //                      + "                           seb.master, "
                //                      + "                           pe.create_time "
                //                     + "              FROM " + schema + ".PACK_ENTRY pe "
                //                      + "                  LEFT JOIN shipping_edi_backup seb "
                //                       + "                    ON pe.pack_id = seb.master "
                //                       + "           WHERE pe.active = 1) "
                //                     + "     WHERE master IS NULL AND (SYSDATE - create_time) * 24 > 72) "
                //              + "  SELECT '" + family + "' as family_id" + " ,ph.psn, "
                //               + "        sc.VALUE AS imei1, "
                //                + "       sc2.VALUE AS imei2, "
                //                + "       aa.pack_id as MPK_Number, "
                //                + "  RW.STATION_GROUP as Current_Stage, "
                //                 + "      aa.create_time as MPK_Conversion_Time, "
                //                 + "      FLOOR ( (SYSDATE - aa.create_time) * 24) AS Exceeded_Hrs "
                //                 + " FROM AA "
                //                  + "     LEFT JOIN " + schema + ".PACK_HISTORY  PH "
                //                   + "       ON AA.PACK_ID = PH.MASTERPACK_NUMBER  "
                //                   + "    LEFT JOIN " + schema + ".SSN_CODE  SC ON PH.PSN = SC.PSN AND SC.NAME = 'IMEI' "
                //                   + "    LEFT JOIN " + schema + ".SSN_CODE  SC2 ON PH.PSN = SC2.PSN AND SC2.NAME = 'IMEI2' "
                //                   + " LEFT JOIN " + schema + ".ROUTING_WIP RW ON SC2.PSN = RW.PSN "
                //                   + " WHERE PH.ACTIVE = 1 ORDER BY aa.pack_id";


                //string strSql = "WITH AA"
                //       + "   AS (SELECT * "
                //        + "     FROM (SELECT DISTINCT "
                //      + "   pe.masterpack_number, "
                //      + "   seb.master, "
                //     + "    pe.PACK_TIME AS create_time "
                //     + "    FROM " + schema + ".PACK_HISTORY pe "
                //     + "    LEFT JOIN shipping_edi_backup seb "
                //     + "       ON pe.masterpack_number = seb.master "
                //      + "   WHERE pe.active = 1 AND pe.PACK_TYPE = 'ATO') "
                //         + "   WHERE master IS NULL AND (SYSDATE - create_time) * 24 > 72) "
                //         + "     SELECT '" + family + "' as family_id" + " , ph.psn, "
                //        + "   sc.VALUE AS imei1, "
                //        + "   sc2.VALUE AS imei2, "
                //        + "   aa.masterpack_number AS MPK_Number, "
                //        + "   RW.STATION_GROUP AS Current_Stage, "
                //       + "    aa.create_time AS MPK_Conversion_Time, "
                //        + "   FLOOR ( (SYSDATE - aa.create_time) * 24) AS Exceeded_Hrs "
                //        + "  FROM AA "
                //       + "    LEFT JOIN " + schema + ".PACK_HISTORY PH "
                //        + "      ON AA.masterpack_number = PH.MASTERPACK_NUMBER "
                //        + "   LEFT JOIN " + schema + ".SSN_CODE SC ON PH.PSN = SC.PSN AND SC.NAME = 'IMEI' "
                //         + "  LEFT JOIN " + schema + ".SSN_CODE SC2 ON PH.PSN = SC2.PSN AND SC2.NAME = 'IMEI2' "
                //         + "  LEFT JOIN " + schema + ".ROUTING_WIP RW ON SC2.PSN = RW.PSN "
                //         + "  WHERE PH.ACTIVE = 1 "
                //         + " ORDER BY aa.masterpack_number ";

                string strSql = "WITH AA "
                                    + " AS (SELECT *  "
                                       + "     FROM (SELECT DISTINCT wo.family_id, "
                                            + "                      wo.model_code, "
                                             + "                     pe.pack_id, "
                                               + "                   seb.master, "
                                               + "                   pe.create_time "
                                               + "     FROM pack_history_h pe "
                                                + "         LEFT JOIN shipping_edi_backup seb "
                                                 + "           ON pe.pack_id = seb.master "
                                                  + "       LEFT JOIN work_order wo "
                                                  + "          ON pe.order_number = wo.order_number "
                                                  + " WHERE     pe.pack_type = 'MASTER' "
                                                   + "      AND pe.pack_level = 'ATO' "
                                                   + "      AND pe.active = 1 "
                                                     + "    AND wo.family_id = '" + family + "' "
                                                     + "    ) "
                                          + " WHERE master IS NULL AND (SYSDATE - create_time) * 24 > 72) "
                                 + " SELECT aa.family_id, "
                                     + "   rw.psn, "
                                     + "   sc.VALUE AS imei1, "
                                      + "  sc2.VALUE AS imei2, "
                                      + "  aa.pack_id AS MPK_Number, "
                                        + " rw.station_group AS Current_Stage, "
                                        + " aa.create_time AS MPK_Conversion_Time, "
                                      + "  FLOOR ( (SYSDATE - aa.create_time) * 24) AS Exceeded_Hrs "
                                   + " FROM AA "
                                     + "   LEFT JOIN " + schema + ".PACK_HISTORY PH ON AA.PACK_ID = PH.MASTERPACK_NUMBER "
                                     + "   LEFT JOIN " + schema + ".SSN_CODE SC ON PH.PSN = SC.PSN AND SC.NAME = 'IMEI' "
                                      + "  LEFT JOIN " + schema + ".SSN_CODE SC2 ON PH.PSN = SC2.PSN AND SC2.NAME = 'IMEI2' "
                                      + "  LEFT JOIN " + schema + ".ROUTING_WIP RW ON SC2.PSN = RW.PSN "
                                   + " WHERE PH.ACTIVE = 1 ";



                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }

        private DataSet GetPDLOutPutData(string family, string schema)
        {
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                string strSql = "WITH aa"
                                    + "   AS (SELECT PSN, STATION_GROUP, last_modified_time "
                                                + " FROM " + schema + ".ROUTING_WIP "
                                                   + " WHERE     PSN IN (SELECT DISTINCT PSN "
                                                     + " FROM " + schema + ".ROUTING_HISTORY "
                                                    + " WHERE station_group = 'PDL' AND STATUS = 'P') "
                                                         + "    AND STATION_GROUP = 'OQC' AND BEFORE_STATION_GROUP='PDL') "
                                                        + "   SELECT '" + family + "' as family_id" + " ,aa.psn, "
                                                        + "  sc.VALUE AS imei, "
                                                         + " aa.station_group AS current_station_group, "
                                                         + " aa.last_modified_time AS pdldatetime "
                                                         + "   FROM aa LEFT JOIN " + schema + ".ssn_code sc ON aa.psn = sc.psn AND sc.name = 'IMEI' "
                                                         + "  WHERE (SYSDATE - aa.LAST_MODIFIED_TIME) * 24 >= 24";
                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }
        private void CollectMFRAutoOutPutData(object sender, System.Timers.ElapsedEventArgs e)
        {
            instockAutoTimer.Stop();
            string timeNow = DateTime.Now.ToString("HHtt");
            try
            {
                if (timeNow == "08AM")
                {
                    WriteLog(timeNow, "Instock_Auto_Method_Executution_Started");
                    DataTable dtMFROutPut = GetInstockData().Tables[0]; //GetMFROutPutData().Tables[0];
                    if (dtMFROutPut.Rows.Count > 0)
                    {
                        DataTable dtMail = GetMailData("INSTOCKInfo").Tables[0];//MFROutPut
                        for (int i = 0; i < dtMail.Rows.Count; i++)
                        {
                            DataRow row = dtMail.Rows[i];
                            Mail ml = new Mail(Mail_User, Mail_Password, Mail_Host, Mail_Port);
                            AddMailAddress(ml, row);
                            ml.Subject = "(FUSE SYSTEM MAIL) " + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "Instock Output Data";
                            ml.Body = "Hi <br/><br/> Please check the attached <b>" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "</b> date Instock Report.<br/><br/><br/><br/>Regards<br/>BharathFIH-IT.";
                            if (AddMFRMailAttachment(ml, dtMFROutPut, "INSTOCKOutPutTemplate", "Data"))
                            {
                                ml.Send2();
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("CollectInstockAutoOutPutData throw exception" + ex.Message, " CollectInstockAutoOutPutData");
                instockAutoTimer.Stop();
                instockAutoTimer.Start();

            }

        }
        protected override void OnStop()
        {
            WriteLog("Service Stopped", " OnStop");
            GC.Collect();
            _runFlag = false;
        }
        private void CollectHourlyInstockOutPutData2()
        {
            WriteLog("Start " + DateTime.Now.ToShortDateString(), "CollectHourlyInstockOutPutData2Method");
            while (_runFlag)
            {
                try
                {
                    int iTaskId = 1010;//Set IND time in DB like 6=>8.30
                    DataTable scheduledTask = this.GetScheduledTask(iTaskId);
                    WriteLog("Called GetScheduledTask " + DateTime.Now.ToShortDateString(), "CollectHourlyInstockOutPutData2Method");
                    bool flag = scheduledTask == null;
                    if (flag)
                    {
                        WriteLog("Wait Schedule " + DateTime.Now.ToShortDateString(), "Schedule");
                        Thread.Sleep(20000);
                        continue;
                    }
                    else
                    {
                        bool flag2 = scheduledTask.Rows.Count == 0;
                        if (flag2)
                        {
                            WriteLog("Wait Schedule " + DateTime.Now.ToShortDateString(), "Schedule");
                            Thread.Sleep(20000);
                            continue;
                        }
                        else
                        {
                            DateTime preDate = Convert.ToDateTime(scheduledTask.Rows[0]["NEXT_TIME"]);
                            int num = Convert.ToInt32(scheduledTask.Rows[0]["TIME_LENGTH"]);
                            //bool flag3 = DateTime.Now.AddMinutes(-10) > preDate.AddMinutes((double)(num));
                            this.MailTriTime = double.Parse(System.Configuration.ConfigurationManager.AppSettings["MailTriggerTime"].ToString());
                            bool flag3 = DateTime.Now.AddMinutes(-MailTriTime) > preDate;
                            if (flag3)
                            {

                                try
                                {
                                    string text2 = string.Empty;
                                    string text3 = string.Empty;
                                    text2 = preDate.ToString("yyyyMMddHHmmss");
                                    DateTime d;
                                    DateTime.TryParseExact(text2, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, out d);

                                    int time = int.Parse(preDate.ToString("HH"));
                                    string timemin = preDate.ToString("HHmmss");
                                    WriteLog("Executed time slot " + time.ToString(), "CollectHourlyInstockOutPutData2Method");
                                    if (time > 6 && time <= 13)
                                    {
                                        text2 = d.ToString("yyyyMMdd083000");
                                        text3 = preDate.ToString("yyyyMMdd") + timemin;
                                        DateTime dd;
                                        DateTime.TryParseExact(text3, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, out dd);
                                        text3 = dd.AddMinutes(+90).ToString("yyyyMMddHHmmss");//chnaged 150 to 90
                                        WriteLog(text2 + "^" + text3, "time_between");

                                    }
                                    else if (time >= 14 && time <= 23)
                                    {
                                        text2 = d.ToString("yyyyMMdd083000");
                                        text3 = preDate.ToString("yyyyMMdd") + timemin;
                                        DateTime dd;
                                        DateTime.TryParseExact(text3, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, out dd);
                                        text3 = dd.AddMinutes(+90).ToString("yyyyMMddHHmmss");//chnaged 150 to 90
                                        WriteLog(text2 + "^" + text3, "time_between");
                                    }
                                    else if (time >= 00 && time <= 06)
                                    {
                                        text2 = d.AddDays(-1).ToString("yyyyMMdd083000");
                                        text3 = preDate.ToString("yyyyMMdd") + timemin;
                                        DateTime dd;
                                        DateTime.TryParseExact(text3, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, out dd);
                                        text3 = dd.AddMinutes(+90).ToString("yyyyMMddHHmmss");//chnaged 150 to 90
                                        WriteLog(text2 + "^" + text3, "time_between");
                                    }
                                    if (CollectInstockHourlyData(text2, text3, preDate))
                                    {
                                        DateTime scheduledTime = preDate.AddMinutes((double)num);
                                        WriteLog("If start " + DateTime.Now.ToShortDateString(), "UpdateSchedualedTaskIF");
                                        this.UpdateSchedualedTask(iTaskId, scheduledTime, "OK", "");
                                        WriteLog("If end " + DateTime.Now.ToShortDateString(), "UpdateSchedualedTaskIF");
                                    }
                                    else
                                    {
                                        DateTime scheduledTime = preDate.AddMinutes((double)num);
                                        WriteLog("Else start " + DateTime.Now.ToShortDateString(), "UpdateSchedualedTaskEnd");
                                        this.UpdateSchedualedTask(iTaskId, scheduledTime, "OK", "");
                                        WriteLog("Else end " + DateTime.Now.ToShortDateString(), "UpdateSchedualedTaskEnd");
                                    }
                                }
                                catch (Exception ex2)
                                {
                                    WriteLog("CollectInstockHourlyData " + ex2.ToString() + " " + DateTime.Now.ToString(), "Exception");
                                }
                            }
                        }
                    }

                }
                catch (Exception ex4)
                {
                    WriteLog(ex4.ToString() + DateTime.Now.ToShortDateString(), "Exception");
                }
                Thread.Sleep(20000);

            }
            WriteLog("End", "CollectHourlyInstockOutPutData...." + DateTime.Now.ToShortDateString());
        }

        private void CollectHourlyInstockOutPutData()
        {
            while (_runFlag)
            {
                try
                {
                    int iTaskId = 1010;//Set IND time in DB like 6=>8.30
                    DataTable scheduledTask = this.GetScheduledTask(iTaskId);
                    bool flag = scheduledTask == null;
                    if (flag)
                    {
                        Thread.Sleep(20000);
                    }
                    else
                    {
                        bool flag2 = scheduledTask.Rows.Count == 0;
                        if (flag2)
                        {
                            Thread.Sleep(20000);
                        }
                        else
                        {
                            DateTime t = Convert.ToDateTime(scheduledTask.Rows[0]["NEXT_TIME"]);
                            //DateTime t = Convert.ToDateTime(scheduledTask.Rows[0]["DB_NOW"]);
                            int num = Convert.ToInt32(scheduledTask.Rows[0]["TIME_LENGTH"]);
                            //bool flag3 = t > dateTime;
                            //bool flag3 = DateTime.Now.AddMinutes((double)(80)) > t;
                            bool flag3 = DateTime.Now.AddMinutes(-10) > t.AddMinutes((double)(num));
                            if (flag3)
                            {
                                try
                                {
                                    string text2 = string.Empty;
                                    string text3 = string.Empty;
                                    text2 = t.ToString("yyyyMMddHHmmss");
                                    //text3 = t.ToString("yyyyMMddHHmmss");
                                    //WriteLog(text2 + "^" + text3, "time_between1");
                                    DateTime d;
                                    DateTime.TryParseExact(text2, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, out d);
                                    //DateTime dateTime;
                                    //DateTime.TryParseExact(text3, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, out dateTime);

                                    //text2 = d.AddMinutes((double)num).ToString("yyyyMMddHHmmss");
                                    //text3 = dateTime.AddMinutes((double)num).ToString("yyyyMMddHHmmss");
                                    //DateTime scheduledTime = t.AddMinutes((double)num);
                                    text2 = d.AddMinutes((double)(150)).ToString("yyyyMMddHHmmss");
                                    text3 = d.AddMinutes((double)(210)).ToString("yyyyMMddHHmmss");
                                    WriteLog(text2 + "^" + text3, "time_between");
                                    if (CollectInstockHourlyOutPutData(text2, text3))
                                    {
                                        DateTime scheduledTime = t.AddMinutes((double)num);
                                        this.UpdateSchedualedTask(iTaskId, scheduledTime, "OK", "");
                                    }


                                }
                                catch (Exception ex2)
                                {
                                    WriteLog(ex2.ToString(), "Exception");
                                }
                            }
                        }
                    }
                }
                catch (Exception ex4)
                {
                    WriteLog(ex4.ToString(), "CollectHourlyInstockOutPutData Exception");
                }
                Thread.Sleep(20000);
            }

        }
        private DataSet GetMailData2(string strGroupType)
        {
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                string strSql = " SELECT * FROM LINE_MONITOR.AUTO_MAIL_GROUP WHERE GROUP_TYPE = '" + strGroupType + "'";
                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }
        private bool CollectInstockHourlyOutPutData(string fdate, string todate)
        {
            bool OutValue = false;
            try
            {
                WriteLog(fdate.ToString(), "Instock_Hourly_Method_Execution_Started");
                DataTable dtMFROutPut = GetMaualInstockData2(fdate, todate).Tables[0]; //GetMFROutPutData().Tables[0];
                if (dtMFROutPut.Rows.Count > 0)
                {
                    DataTable dtMail = GetMailData2("INSTOCKHInfo").Tables[0];//MFROutPut
                    for (int i = 0; i < dtMail.Rows.Count; i++)
                    {
                        //string onlyDate =Convert.ToDateTime(fdate).ToShortDateString();
                        DateTime d;
                        DateTime.TryParseExact(fdate, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, out d);
                        DateTime d2;
                        DateTime.TryParseExact(todate, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, out d2);

                        string fromDate = d.AddMinutes((double)(-150)).ToString("HH:mm tt");
                        string toDate = d2.AddMinutes((double)(-150)).ToString("HH:mm tt");
                        string subject = d.ToString("yyyy-MM-dd") + " " + fromDate + " to " + toDate;
                        DataRow row = dtMail.Rows[i];
                        Mail ml = new Mail(Mail_User, Mail_Password, Mail_Host, Mail_Port);
                        AddMailAddress(ml, row);
                        ml.Subject = "(FUSE SYSTEM MAIL) " + subject + " Instock Hourly Output Data";
                        ml.Body = "Hi <br/><br/> Please check the attached <b>" + subject + " </b> datetime Instock Report.<br/><br/><br/><br/>Regards<br/>BharathFIH-IT.";
                        if (AddMFRMailAttachment2(ml, dtMFROutPut, "INSTOCKOutPutTemplate", "Data", subject))
                        {
                            ml.Send2();

                            return OutValue = true;
                        }
                        else
                        {
                            OutValue = false;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("Instock_Hourly_Method_Execution_Started throw exception:" + ex.Message, " CollectInstockHourlyOutPutData");
            }

            return OutValue;
        }
        private int UpdateSchedualedTask(int iTaskId, DateTime scheduledTime, string strStatus, string strMessage)
        {
            int result = 0;
            string text = " update sfis1.SCHEDULED_TASKS set EXECUTE_TIME=sysdate,NEXT_TIME=:NEXT_TIME,STATUS=:STATUS,REPORT_MSG=:REPORT_MSG where TASK_ID=:TASK_ID ";
            try
            {
                using (OracleConnection oracleConnection = new OracleConnection(this._DBConnStr))
                {
                    oracleConnection.Open();
                    result = new OracleCommand(text.ToString(), oracleConnection)
                    {
                        Parameters = 
						{
							new OracleParameter("NEXT_TIME", scheduledTime),
							new OracleParameter("STATUS", strStatus),
							new OracleParameter("REPORT_MSG", strMessage),
							new OracleParameter("TASK_ID", iTaskId)
						}
                    }.ExecuteNonQuery();
                }
            }
            catch (Exception ex1)
            {
                WriteLog(ex1.ToString(), "UpdateSchedualedTask");
                result = 0;
            }
            return result;
        }

        private DataTable GetScheduledTask(int iTaskId)
        {
            string text = " Select a.*,sysdate db_now From sfis1.SCHEDULED_TASKS a where TASK_ID=:TASK_ID ";
            DataTable result;
            try
            {
                using (OracleConnection oracleConnection = new OracleConnection(this._DBConnStr))
                {
                    oracleConnection.Open();
                    OracleDataAdapter oracleDataAdapter = new OracleDataAdapter(new OracleCommand(text.ToString(), oracleConnection)
                    {
                        Parameters = 
						{
							new OracleParameter("TASK_ID", iTaskId)
						}
                    });
                    DataSet dataSet = new DataSet();
                    oracleDataAdapter.Fill(dataSet);
                    result = dataSet.Tables[0];
                }
            }
            catch (Exception ex5)
            {
                WriteLog(ex5.ToString(), "GetScheduledTask");
                result = null;
            }
            return result;
        }
        private void CollectMFRManualOutPutData(string fdate, string todate)
        {
            while (_runFlag)
            {
                //int iHour = DateTime.Now.Hour;
                //int iMinute = DateTime.Now.Minute;
                //if (iHour == 08 && iMinute == 00)
                //{
                WriteLog(fdate.ToString(), "Instock_Manual_Method_Execution_Started");
                try
                {
                    DataTable dtMFROutPut = GetMaualInstockData(fdate, todate).Tables[0]; //GetMFROutPutData().Tables[0];
                    if (dtMFROutPut.Rows.Count > 0)
                    {
                        DataTable dtMail = GetMailData("INSTOCKInfo").Tables[0];//MFROutPut
                        for (int i = 0; i < dtMail.Rows.Count; i++)
                        {
                            string fromDate = Convert.ToDateTime(fdate).ToString("yyyy-MM-dd");
                            DataRow row = dtMail.Rows[i];
                            Mail ml = new Mail(Mail_User, Mail_Password, Mail_Host, Mail_Port);
                            AddMailAddress(ml, row);
                            ml.Subject = "(FUSE SYSTEM MAIL) " + fromDate + "Instock Output Data";
                            ml.Body = "Hi <br/><br/> Please check the attached <b>" + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + "</b> date Instock Report.<br/><br/><br/><br/>Regards<br/>BharathFIH-IT.";
                            if (AddMFRMailAttachment(ml, dtMFROutPut, "INSTOCKOutPutTemplate", "Data", fdate))
                            {
                                ml.Send2();
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    WriteLog("CollectInstockManualOutPutData throw exception:" + ex.Message, " CollectInstockManualOutPutData");
                }
                //}
                Thread.Sleep(60000);
                _runFlag = false;
            }
        }
        private void CollectMFRAutoOutPutData()
        {
            WriteLog("Start.. CollectInstockAutoOutPutData Method " + DateTime.Now.ToString(), "CollectInstockAutoOutPutData");
            while (_runFlag)
            {
                //int dd = DateTime.Now.Hour;
                //int mm = DateTime.Now.Minute;
                int iHour = DateTime.Now.Hour;//8; //;//InstockAutoMail.Default.IsMin;
                int iMinute = DateTime.Now.Minute;// 10; //DateTime.Now.Minute;// InstockAutoMail.Default.IsSec;// 
                string timeNow2 = DateTime.Now.ToString("HH:mmtt");
                string ObatimeNow = ConfigurationManager.AppSettings["MailTriggerOBATime"].ToString();
                //DateTime.Now.Hour.ToString() == iHour && DateTime.Now.Minute.ToString() == iMinute
                int iHours = int.Parse(InstockAutoMail.Default.IsHour);
                int iMinutes = int.Parse(InstockAutoMail.Default.IsMin);
                string eTime = iHour + ":" + iMinute;
                if (timeNow2 == ObatimeNow)//iHour == iHours && iMinute == iMinutes
                {
                    WriteLog(eTime, "CollectInstockAutoOutPutData");
                    try
                    {
                        DataTable dtMFROutPut = GetInstockData().Tables[0]; //GetMFROutPutData().Tables[0];
                        WriteLog("Get Instock Data " + DateTime.Now.ToString(), "CollectInstockAutoOutPutData");
                        if (dtMFROutPut.Rows.Count > 0)
                        {
                            DataTable dtMail = GetMailData("INSTOCKInfo").Tables[0];//MFROutPut
                            WriteLog("Call Mail Id's " + DateTime.Now.ToString(), "CollectInstockAutoOutPutData");
                            for (int i = 0; i < dtMail.Rows.Count; i++)
                            {
                                DataRow row = dtMail.Rows[i];
                                Mail ml = new Mail(Mail_User, Mail_Password, Mail_Host, Mail_Port);
                                AddMailAddress(ml, row);
                                ml.Subject = "(FUSE SYSTEM MAIL) " + DateTime.Now.AddDays(-1).ToString("yyyy MMMM dd") + "Instock Output Data";
                                ml.Body = "Hi <br/><br/> Please check the attached <b>" + DateTime.Now.AddDays(-1).ToString("yyyy MMMM dd") + "</b> date Instock Report.<br/><br/><br/><br/>Regards<br/>BharathFIH-IT.";
                                if (AddMFRMailAttachment(ml, dtMFROutPut, "INSTOCKOutPutTemplate", "Data"))
                                {
                                    ml.Send2();
                                    WriteLog("Sent file to Mail Id's " + DateTime.Now.ToString(), "CollectInstockAutoOutPutData");
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteLog("CollectInstockAutoOutPutData throw exception" + ex.Message, " CollectInstockAutoOutPutData");
                    }
                }
                Thread.Sleep(60000);
            }
            WriteLog("End.. CollectInstockAutoOutPutData Method " + DateTime.Now.ToString(), "CollectInstockAutoOutPutData");
        }
        private DataSet GetInstockData()
        {
            string p_time_start = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 08:30");
            string p_time_end = DateTime.Now.ToString("yyyy-MM-dd 08:30");
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                string strSql = "SELECT project,group_level,product_id AS MODEL_CODE,SUM (quan) qty FROM B2B.SRM_PRODUCT@insrm01 sp " +
                           " WHERE group_level = 'INSTOCK' AND begin_time >= TO_DATE ('" + p_time_start + "', 'YYYY-MM-DD HH24:MI:SS') " +
                           "AND begin_time < TO_DATE ('" + p_time_end + "', 'YYYY-MM-DD HH24:MI:SS') GROUP BY project, group_level, product_id " +
                           "ORDER BY product_id";
                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }
        //Manual
        private DataSet GetMaualInstockData(string from, string to)
        {
            string p_time_start = Convert.ToDateTime(from).ToString("yyyy-MM-dd 08:30");
            string p_time_end = Convert.ToDateTime(to).ToString("yyyy-MM-dd 08:30");
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                string strSql = "SELECT project,group_level,product_id AS MODEL_CODE,SUM (quan) qty FROM B2B.SRM_PRODUCT@insrm01 sp " +
                           " WHERE group_level = 'INSTOCK' AND begin_time >= TO_DATE ('" + p_time_start + "', 'YYYY-MM-DD HH24:MI:SS') " +
                           "AND begin_time < TO_DATE ('" + p_time_end + "', 'YYYY-MM-DD HH24:MI:SS') GROUP BY project, group_level, product_id " +
                           "ORDER BY product_id";
                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }
        private DataSet GetMaualInstockData2(string from, string to)
        {
            //string p_time_start = Convert.ToDateTime(from).AddHours(+2.30).ToString("yyyy-MM-dd HH:mm:00"); //Convert.ToDateTime(from).ToString("yyyy-MM-dd 08:30");
            //string p_time_end = Convert.ToDateTime(to).AddHours(+2.30).ToString("yyyy-MM-dd HH:mm:00");//Convert.ToDateTime(to).ToString("yyyy-MM-dd 08:30");
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                string strSql = "SELECT project,group_level,product_id AS MODEL_CODE,SUM (quan) qty FROM B2B.SRM_PRODUCT@insrm01 sp " +
                           " WHERE group_level = 'INSTOCK' AND begin_time >= TO_DATE ('" + from + "', 'YYYY-MM-DD HH24:MI:SS') " +
                           "AND begin_time < TO_DATE ('" + to + "', 'YYYY-MM-DD HH24:MI:SS') GROUP BY project, group_level, product_id " +
                           "ORDER BY product_id";
                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }

        private bool CollectInstockHourlyData(string rfrdate, string rtdate, DateTime rtime)
        {
            bool rovalue = false;
            WriteLog("Fetch Instock Data Start... " + DateTime.Now.ToShortDateString(), "InstockFetchMethod");
            DataTable dtMFROutPut = GetMaualInstockData5(rfrdate, rtdate).Tables[0]; //GetMFROutPutData().Tables[0];
            WriteLog("Fetch Instock Data End...  " + dtMFROutPut.Rows.Count.ToString() + " Records (" + DateTime.Now.ToShortDateString() + ")", "InstockFetchMethod");
            DataRow addRow = dtMFROutPut.NewRow();
            DataColumn a = new DataColumn("ATotal", typeof(decimal));
            dtMFROutPut.Columns.Add(a);
            DataColumn b = new DataColumn("BTotal", typeof(decimal));
            dtMFROutPut.Columns.Add(b);
            DataColumn c = new DataColumn("CTotal", typeof(decimal));
            dtMFROutPut.Columns.Add(c);
            DataColumn e = new DataColumn("DTime", typeof(string));
            dtMFROutPut.Columns.Add(e);
            DataColumn f = new DataColumn("GrandTotal", typeof(decimal));
            dtMFROutPut.Columns.Add(f);
            DateTime dd;
            DateTime.TryParseExact(rfrdate, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, out dd);

            if (dtMFROutPut.Rows.Count > 0)
            {
                foreach (DataRow item in dtMFROutPut.Rows)
                {
                    item["DTime"] = dd.ToString("yyyy/MM/dd");
                    item["ATotal"] = int.Parse(item["B1"].ToString()) + int.Parse(item["B2"].ToString()) + int.Parse(item["B3"].ToString()) + int.Parse(item["B4"].ToString()) + int.Parse(item["B5"].ToString()) + int.Parse(item["B6"].ToString()) + int.Parse(item["B7"].ToString()) + int.Parse(item["B8"].ToString());
                    item["BTotal"] = int.Parse(item["B9"].ToString()) + int.Parse(item["B10"].ToString()) + int.Parse(item["B11"].ToString()) + int.Parse(item["B12"].ToString()) + int.Parse(item["B13"].ToString()) + int.Parse(item["B14"].ToString()) + int.Parse(item["B15"].ToString()) + int.Parse(item["B16"].ToString());
                    item["CTotal"] = int.Parse(item["B17"].ToString()) + int.Parse(item["B18"].ToString()) + int.Parse(item["B19"].ToString()) + int.Parse(item["B20"].ToString()) + int.Parse(item["B21"].ToString()) + int.Parse(item["B22"].ToString()) + int.Parse(item["B23"].ToString()) + int.Parse(item["B24"].ToString());
                    item["GrandTotal"] = int.Parse(item["ATotal"].ToString()) + int.Parse(item["BTotal"].ToString()) + int.Parse(item["CTotal"].ToString());
                }

                double aTotalSum = Convert.ToDouble(dtMFROutPut.Compute("SUM(ATotal)", string.Empty));
                double bTotalSum = Convert.ToDouble(dtMFROutPut.Compute("SUM(BTotal)", string.Empty));
                double cTotalSum = Convert.ToDouble(dtMFROutPut.Compute("SUM(CTotal)", string.Empty));
                double abcTotalSum = Convert.ToDouble(dtMFROutPut.Compute("SUM(GrandTotal)", string.Empty));
                addRow["ATotal"] = aTotalSum;
                addRow["BTotal"] = bTotalSum;
                addRow["CTotal"] = cTotalSum;
                addRow["GrandTotal"] = abcTotalSum;
                dtMFROutPut.Rows.Add(addRow);

                DataTable dtMail = GetMailData2("INSTOCKHInfo").Tables[0];//MFROutPut
                for (int i = 0; i < dtMail.Rows.Count; i++)
                {
                    DataRow row = dtMail.Rows[i];
                    Mail ml = new Mail(Mail_User, Mail_Password, Mail_Host, Mail_Port);
                    AddMailAddress(ml, row);
                    DateTime d;
                    DateTime.TryParseExact(rfrdate, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, out d);
                    string subject = d.ToString("yyyy-MM-dd") + " " + rtime.AddMinutes(-60).ToString("HH") + " to " + rtime.ToString("HH");
                    string d2 = d.ToString("yyyy-MM-dd");
                    ml.Subject = "(FUSE SYSTEM MAIL)  Instock " + subject + " Hourly Output Data";
                    ml.Body = "Hi <br/><br/> Please check the attached " + subject + "  </b> datetime Instock Report.<br/><br/><br/><br/>Regards<br/>" + _CName + ".";
                    if (AddMFRMailAttachment3(ml, dtMFROutPut, "INSTOCKHourlyOutPutTemplate", "Data", "INSTOCKHourlyData " + subject))
                    {

                        try
                        {
                            WriteLog("Mail Start... " + DateTime.Now.ToShortDateString(), "MailMethod");
                            ml.Send2();
                            rovalue = true;
                            WriteLog("Mail End... " + DateTime.Now.ToShortDateString(), "MailMethod");
                        }
                        catch (Exception ex)
                        {
                            WriteLog("Mail error message... " + DateTime.Now.ToShortDateString() + " " + ex.Message, "MailMethod");
                        }

                    }
                    else
                    {
                        break;
                    }
                }
            }
            //else
            //{
            //    rovalue = true;
            //}
            return rovalue;
        }
        private DataSet GetMaualInstockData4(string fdate, string tdate)
        {
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                string strSql = "select * from ( SELECT project," +
                                                   "to_char (begin_time, 'HH24') AS dtime," +
                                                   "group_level," +
                                                   "product_id AS MODEL_CODE," +
                                                   "SUM (quan) qty " +
                                              "FROM B2B.SRM_PRODUCT@INSRM01 sp" +
                                            " WHERE     group_level = 'INSTOCK'" +
                                                 "  AND begin_time >=" +
                                                      "    TO_DATE (" + fdate + "," +
                                                                   "'YYYY-MM-DD HH24:MI:SS')" +
                                                   " AND begin_time <" +
                                                      "    TO_DATE (" + tdate + "," +
                                                                   "'YYYY-MM-DD HH24:MI:SS')" +
                                          "GROUP BY project," +
                                            "to_char (begin_time, 'HH24')," +
                                                "group_level," +
                                                "product_id) a " +
                                "pivot(" +
                                "sum(qty) for dtime in" +
                                "('08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23'," +
                                "'00','01','02','03','04','05','06','07')) " +
                                "order by project";
                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }

        private DataSet GetMaualInstockData5(string fdate, string tdate)
        {
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                string strSql = "select project,"
                                    + "group_level,"
                                    + "MODEL_CODE,"
                                    + "NVL (a1, 0) AS b1,"
                                    + "NVL (a2, 0) AS b2,"
                                    + "NVL (a3, 0) AS b3,"
                                    + "NVL (a4, 0) AS b4,"
                                    + "NVL (a5, 0) AS b5,"
                                    + "NVL (a6, 0) AS b6,"
                                    + "NVL (a7, 0) AS b7,"
                                    + "NVL (a8, 0) AS b8,"
                                    + "NVL (a9, 0) AS b9,"
                                    + "NVL (a10, 0) AS b10,"
                                    + "NVL (a11, 0) AS b11,"
                                    + "NVL (a12, 0) AS b12,"
                                    + "NVL (a13, 0) AS b13,"
                                    + "NVL (a14, 0) AS b14,"
                                    + "NVL (a15, 0) AS b15,"
                                    + "NVL (a16, 0) AS b16,"
                                    + "NVL (a17, 0) AS b17,"
                                    + "NVL (a18, 0) AS b18,"
                                    + "NVL (a19, 0) AS b19,"
                                    + "NVL (a20, 0) AS b20,"
                                    + "NVL (a21, 0) AS b21,"
                                    + "NVL (a22, 0) AS b22,"
                                    + "NVL (a23, 0) AS b23,"
                                    + "NVL (a24, 0) AS b24 " +
                                                  "from ( SELECT project," +
                                                   "to_char (begin_time, 'HH24') AS dtime," +
                                                   "group_level," +
                                                   "product_id AS MODEL_CODE," +
                                                   "SUM (quan) qty " +
                                              "FROM B2B.SRM_PRODUCT@INSRM01 sp" +
                                            " WHERE     group_level = 'INSTOCK'" +
                                                 "  AND begin_time >=" +
                                                      "    TO_DATE (" + fdate + "," +
                                                                   "'YYYY-MM-DD HH24:MI:SS')" +
                                                   " AND begin_time <=" +
                                                      "    TO_DATE (" + tdate + "," +
                                                                   "'YYYY-MM-DD HH24:MI:SS')" +
                                          "GROUP BY project," +
                                            "to_char (begin_time, 'HH24')," +
                                                "group_level," +
                                                "product_id) a " +
                                "pivot(" +
                                "sum(qty) for dtime in " +
                                "('08' AS a1,'09' AS a2,'10' AS a3,'11' AS a4,'12' AS a5,'13' AS a6,'14' AS a7,'15' AS a8,"
                                     + "'16' AS a9,'17' AS a10,'18' AS a11,'19' AS a12,'20' AS a13,'21' AS a14,'22' AS a15,'23' AS a16,"
                                     + "'00' AS a17,'01' AS a18,'02' AS a19,'03' AS a20,'04' AS a21,'05' AS a22,'06' AS a23,'07' AS a24"
                                     + ")) "
                                + "order by project";
                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }

        private DataSet GetMaualInstockData3(string fdate, string tdate)
        {
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                string strSql = "WITH bb (dd, dd1)" +
                                "AS (SELECT TO_DATE (" + fdate + ", 'YYYY-MM-DD HH24:MI:SS')" +
                                 "AS dd," +
                                  "TO_DATE (" + fdate + ", 'YYYY-MM-DD HH24:MI:SS')" +
                                  "+ 1 / 24" +
                                 "AS dd1 " +
                                "FROM DUAL " +
                                 "UNION ALL " +
                                 "SELECT dd1 AS dd, dd1 + 1 / 24 AS dd1" +
                                " FROM bb " +
                                 "WHERE dd1 <" +
                                 " TO_DATE (" + tdate + ", 'YYYY-MM-DD HH24:MI:SS')) " +
                                "select * from (" +
                                        " SELECT a.begin_time,a.project AS project," +
                                         "TO_CHAR (bb.dd-2, 'HH24')||'-'||TO_CHAR (bb.dd1-2, 'HH24') as TIME," +
                                         "a.group_level," +
                                         "a.MODEL_CODE," +
                                         "NVL (a.qty, '0') AS qty " +
                                    "FROM bb " +
                                         "LEFT JOIN" +
                                         "(  SELECT project," +
                                                   "begin_time," +
                                                   "end_time," +
                                                   "group_level," +
                                                   "product_id AS MODEL_CODE," +
                                                   "SUM (quan) qty " +
                                              "FROM B2B.SRM_PRODUCT@INSRM01 sp" +
                                            " WHERE     group_level = 'INSTOCK'" +
                                                 "  AND begin_time >=" +
                                                      "    TO_DATE (" + fdate + "," +
                                                                   "'YYYY-MM-DD HH24:MI:SS')" +
                                                   " AND begin_time <" +
                                                      "    TO_DATE (" + tdate + "," +
                                                                   "'YYYY-MM-DD HH24:MI:SS')" +
                                          "GROUP BY project," +
                                            "begin_time," +
                                             "end_time," +
                                                "group_level," +
                                                "product_id) a " +
                                          "  ON bb.dd = a.begin_time AND bb.dd1 = a.end_time " +
                                "ORDER BY TO_CHAR (bb.dd, 'YYYY-MM-DD HH24:MI:SS'))a " +
                                "pivot(" +
                                "sum(qty) for time in" +
                                "('08-09','09-10','10-11','11-12','12-13','13-14','14-15','15-16','16-17','17-18','18-19','19-20','20-21','21-22','22-23','23-00'," +
                                "'00-01','01-02','02-03','03-04','04-05','05-06','06-07','07-08')) " +
                                "order by project";
                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }


        private DataSet GetMailData(string strGroupType)
        {
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                string strSql = " SELECT * FROM LINE_MONITOR.AUTO_MAIL_GROUP WHERE GROUP_TYPE = '" + strGroupType + "'";
                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }
        private DataSet GetMailDataPdl(string strGroupType)
        {
            DataSet myds = new DataSet();
            using (OracleConnection orcn = new OracleConnection(this._DBConnStr))
            {
                string strSql = " SELECT * FROM LINE_MONITOR.AUTO_MAIL_GROUP WHERE GROUP_TYPE = '" + strGroupType + "'";
                OracleCommand cmd = new OracleCommand(strSql, orcn);
                try
                {
                    orcn.Open();
                    OracleDataAdapter oda = new OracleDataAdapter(cmd);
                    oda.Fill(myds);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMailData, error message: " + ex.Message, " Public");
                }
            }
            return myds;
        }
        private void AddMailAddress(Mail ml, DataRow row)
        {
            string[] mailToArray = row["MAIL_TO"].ToString().Trim(';').Split(';');
            string[] mailCcArray = row["MAIL_CC"].ToString().Trim(';').Split(';');
            string[] mailBCcArray = row["MAIL_BCC"].ToString().Trim(';').Split(';');
            //int count = int.Parse(System.Configuration.ConfigurationManager.AppSettings["CC"].ToString());
            //mailCcArray = mailCcArray.Take(count).ToArray();
            mailCcArray = mailCcArray.ToArray();

            for (int i = 0; i < mailToArray.Length; i++)
            {
                ml.To.Add(mailToArray[i]);
            }
            for (int i = 0; i < mailCcArray.Length; i++)
            {
                ml.Cc.Add(mailCcArray[i]);
            }

            for (int i = 0; i < mailBCcArray.Length; i++)
            {
                ml.Bcc.Add(mailBCcArray[i]);
            }

            ml.From = Mail_User + "@foxconn.com";
            byte[] img = new byte[0];
            string strBody = "";
            ml.Body = strBody;
        }

        private bool AddVUTVMailAttachment(Mail ml, DataTable dt, string strTemplateName, string strSheetName)
        {
            string fname = "VUTV";
            try
            {
                string FileDirectory = serviceInstallPath + @"files\VUTV";
                if (!Directory.Exists(FileDirectory))
                {
                    Directory.CreateDirectory(FileDirectory);
                }
                if (dt.Rows.Count > 0)
                {
                    string filename = fname + "_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
                    string filePath = FileDirectory + @"\" + filename;
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }
                    //傳入模板名稱及sheet名稱
                    //FUSE.ExportExcel.TableToExcel(dt, @"D:\MFROutPutTemplate.xls");
                    //TableToExcelForXLSX(dt, filePath); 
                    MemoryStream ms = FUSE.ExportExcel.GeneratePublicSend(dt, strTemplateName, strSheetName);
                    StreamWriter sw = new StreamWriter(filePath, true, System.Text.Encoding.GetEncoding("GB2312"));

                    ms.WriteTo(sw.BaseStream);
                    sw.Close();
                    ml.Attachment.Add(filePath);
                }
                return true;
            }
            catch (Exception ex)
            {
                WriteLog("Error accurs on method: AddMailBody, error message: " + ex.Message, " CollectInstockOutPutData");
                return false;
            }
        }

        public void TableToExcelForXLSX(DataTable dt, string file)
        {
            XSSFWorkbook xssfworkbook = new XSSFWorkbook();
            ISheet sheet = xssfworkbook.CreateSheet("Data");

            //表头
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }

            //数据
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row1 = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = row1.CreateCell(j);
                    cell.SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            //转为字节数组
            MemoryStream stream = new MemoryStream();
            xssfworkbook.Write(stream);
            var buf = stream.ToArray();

            //保存为Excel文件
            using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                fs.Write(buf, 0, buf.Length);
                fs.Flush();
            }
        }




        private bool AddMFRMailAttachment(Mail ml, DataTable dt, string strTemplateName, string strSheetName)
        {
            try
            {
                string FileDirectory = serviceInstallPath + @"files\Instock";
                if (!Directory.Exists(FileDirectory))
                {
                    Directory.CreateDirectory(FileDirectory);
                }
                if (dt.Rows.Count > 0)
                {
                    string filename = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") + ".xls";
                    string filePath = FileDirectory + @"\" + filename;
                    if (File.Exists(filePath))
                    {
                        return false;
                        //File.Delete(filePath);
                    }
                    //傳入模板名稱及sheet名稱
                    //FUSE.ExportExcel.TableToExcel(dt, @"D:\MFROutPutTemplate.xls");
                    MemoryStream ms = FUSE.ExportExcel.GeneratePublicSend(dt, strTemplateName, strSheetName);
                    StreamWriter sw = new StreamWriter(filePath, true, System.Text.Encoding.GetEncoding("GB2312"));
                    ms.WriteTo(sw.BaseStream);
                    sw.Close();
                    ml.Attachment.Add(filePath);
                }
                return true;
            }
            catch (Exception ex)
            {
                WriteLog("Error accurs on method: AddMailBody, error message: " + ex.Message, " CollectInstockOutPutData");
                return false;
            }
        }

        private bool AddMPKINSTOCKGAPMailAttachment(Mail ml, DataTable dt, string strTemplateName, string strSheetName)
        {
            string fname = "Mpk_Instock_Gap_";
            try
            {
                string FileDirectory = serviceInstallPath + @"files\MPKINSTOCK";
                if (!Directory.Exists(FileDirectory))
                {
                    Directory.CreateDirectory(FileDirectory);
                }
                if (dt.Rows.Count > 0)
                {
                    string filename = fname + " " + DateTime.Now.ToString("yyyy-MM-dd HHmmtt") + ".xls";
                    string filePath = FileDirectory + @"\" + filename;
                    if (File.Exists(filePath))
                    {
                        return false;
                        //File.Delete(filePath);
                    }
                    //傳入模板名稱及sheet名稱
                    //FUSE.ExportExcel.TableToExcel(dt, @"D:\MFROutPutTemplate.xls");
                    MemoryStream ms = FUSE.ExportExcel.GeneratePublicSend(dt, strTemplateName, strSheetName);
                    StreamWriter sw = new StreamWriter(filePath, true, System.Text.Encoding.GetEncoding("GB2312"));
                    ms.WriteTo(sw.BaseStream);
                    sw.Close();
                    ml.Attachment.Add(filePath);
                }
                return true;
            }
            catch (Exception ex)
            {
                WriteLog("Error accurs on method: AddMailBody, error message: " + ex.Message, " CollectMpkInstockGapData");
                return false;
            }
        }


        private bool AddOQCDISPAYFAILMailAttachment(Mail ml, DataTable dt, string strTemplateName, string strSheetName)
        {
            string fname = "OQC_Disp_Fail_";
            try
            {
                string FileDirectory = serviceInstallPath + @"files\OQCDipFail";
                if (!Directory.Exists(FileDirectory))
                {
                    Directory.CreateDirectory(FileDirectory);
                }
                if (dt.Rows.Count > 0)
                {
                    string filename = fname + " " + DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd HHmmtt") + ".xls";
                    string filePath = FileDirectory + @"\" + filename;
                    if (File.Exists(filePath))
                    {
                        return false;
                        //File.Delete(filePath);
                    }
                    //傳入模板名稱及sheet名稱
                    //FUSE.ExportExcel.TableToExcel(dt, @"D:\MFROutPutTemplate.xls");
                    MemoryStream ms = FUSE.ExportExcel.GeneratePublicSend(dt, strTemplateName, strSheetName);
                    StreamWriter sw = new StreamWriter(filePath, true, System.Text.Encoding.GetEncoding("GB2312"));
                    ms.WriteTo(sw.BaseStream);
                    sw.Close();
                    ml.Attachment.Add(filePath);
                }
                return true;
            }
            catch (Exception ex)
            {
                WriteLog("Error accurs on method: AddMailBody, error message: " + ex.Message, " CollectInstockOutPutData");
                return false;
            }
        }
        private bool AddOBAMailAttachment(Mail ml, DataTable dt, string strTemplateName, string strSheetName)
        {
            string fname = dt.Rows[0][0].ToString();
            try
            {
                string FileDirectory = serviceInstallPath + @"files\OBA";
                if (!Directory.Exists(FileDirectory))
                {
                    Directory.CreateDirectory(FileDirectory);
                }
                if (dt.Rows.Count > 0)
                {
                    string filename = fname + " " + DateTime.Now.ToString("yyyy-MM-dd HHmmtt") + ".xls";
                    string filePath = FileDirectory + @"\" + filename;
                    if (File.Exists(filePath))
                    {
                        return false;
                        //File.Delete(filePath);
                    }
                    //傳入模板名稱及sheet名稱
                    //FUSE.ExportExcel.TableToExcel(dt, @"D:\MFROutPutTemplate.xls");
                    MemoryStream ms = FUSE.ExportExcel.GeneratePublicSend(dt, strTemplateName, strSheetName);
                    StreamWriter sw = new StreamWriter(filePath, true, System.Text.Encoding.GetEncoding("GB2312"));
                    ms.WriteTo(sw.BaseStream);
                    sw.Close();
                    ml.Attachment.Add(filePath);
                }
                return true;
            }
            catch (Exception ex)
            {
                WriteLog("Error accurs on method: AddMailBody, error message: " + ex.Message, " CollectInstockOutPutData");
                return false;
            }
        }
        private bool AddMPKShippedMailAttachment(Mail ml, DataTable dt, string strTemplateName, string strSheetName)
        {
            string fname = dt.Rows[0][0].ToString();
            try
            {
                string FileDirectory = serviceInstallPath + @"files\MPKShipped";
                if (!Directory.Exists(FileDirectory))
                {
                    Directory.CreateDirectory(FileDirectory);
                }
                if (dt.Rows.Count > 0)
                {
                    string filename = fname + " " + DateTime.Now.ToString("yyyy-MM-dd HHmmtt") + ".xls";
                    string filePath = FileDirectory + @"\" + filename;
                    if (File.Exists(filePath))
                    {
                        return false;
                        //File.Delete(filePath);
                    }
                    //傳入模板名稱及sheet名稱
                    //FUSE.ExportExcel.TableToExcel(dt, @"D:\MFROutPutTemplate.xls");
                    // MemoryStream ms = FUSE.ExportExcel.GeneratePublicSend(dt, strTemplateName, strSheetName);
                    MemoryStream ms = FUSE.ExportExcel.GeneratePublicSend(dt, strTemplateName, strSheetName);
                    StreamWriter sw = new StreamWriter(filePath, true, System.Text.Encoding.GetEncoding("GB2312"));
                    ms.WriteTo(sw.BaseStream);
                    sw.Close();
                    ml.Attachment.Add(filePath);
                }
                return true;
            }
            catch (Exception ex)
            {
                WriteLog("Error accurs on method: AddMailBody, error message: " + ex.Message, " CollectMPKShippedAutoOutPutData");
                return false;
            }
        }
        private bool AddOBAMailAttachment2(Mail ml, DataTable dt, string strTemplateName, string strSheetName)
        {
            string fname = dt.Rows[0][0].ToString();
            try
            {
                string FileDirectory = serviceInstallPath + @"files\OBAPDL";
                if (!Directory.Exists(FileDirectory))
                {
                    Directory.CreateDirectory(FileDirectory);
                }
                if (dt.Rows.Count > 0)
                {
                    string filename = fname + " " + DateTime.Now.ToString("yyyy-MM-dd HHmmtt") + ".xls";
                    string filePath = FileDirectory + @"\" + filename;
                    if (File.Exists(filePath))
                    {
                        return false;
                        //File.Delete(filePath);
                    }
                    //傳入模板名稱及sheet名稱
                    //FUSE.ExportExcel.TableToExcel(dt, @"D:\MFROutPutTemplate.xls");
                    MemoryStream ms = FUSE.ExportExcel.GeneratePublicSend(dt, strTemplateName, strSheetName);
                    StreamWriter sw = new StreamWriter(filePath, true, System.Text.Encoding.GetEncoding("GB2312"));
                    ms.WriteTo(sw.BaseStream);
                    sw.Close();
                    ml.Attachment.Add(filePath);
                }
                return true;
            }
            catch (Exception ex)
            {
                WriteLog("Error accurs on method: AddMailBody, error message: " + ex.Message, " CollectInstockOutPutData");
                return false;
            }
        }
        //Manual
        private bool AddMFRMailAttachment(Mail ml, DataTable dt, string strTemplateName, string strSheetName, string fdate)
        {
            try
            {
                string FileDirectory = serviceInstallPath + @"files\Instock";
                if (!Directory.Exists(FileDirectory))
                {
                    Directory.CreateDirectory(FileDirectory);
                }
                if (dt.Rows.Count > 0)
                {
                    string filename = Convert.ToDateTime(fdate).ToString("yyyy-MM-dd") + ".xls";
                    string filePath = FileDirectory + @"\" + filename;
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }
                    //傳入模板名稱及sheet名稱
                    //FUSE.ExportExcel.TableToExcel(dt, @"D:\MFROutPutTemplate.xls");
                    MemoryStream ms = FUSE.ExportExcel.GeneratePublicSend(dt, strTemplateName, strSheetName);
                    StreamWriter sw = new StreamWriter(filePath, true, System.Text.Encoding.GetEncoding("GB2312"));
                    ms.WriteTo(sw.BaseStream);
                    sw.Close();
                    ml.Attachment.Add(filePath);
                }
                return true;
            }
            catch (Exception ex)
            {
                WriteLog("Error accurs on method: AddMailBody, error message: " + ex.Message, " CollectInstockOutPutData");
                return false;
            }
        }
        private bool AddMFRMailAttachment2(Mail ml, DataTable dt, string strTemplateName, string strSheetName, string sub)
        {
            try
            {
                string FileDirectory = serviceInstallPath + @"files\InstockHourly\";
                if (!Directory.Exists(FileDirectory))
                {
                    Directory.CreateDirectory(FileDirectory);
                }
                if (dt.Rows.Count > 0)
                {
                    //DateTime d;
                    //DateTime.TryParseExact(fdate, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, out d);
                    //DateTime d2;
                    //DateTime.TryParseExact(tdate, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, out d2);

                    //string fromDate = d.AddMinutes((double)(-150)).ToString("yyyy-MM-dd HHmm");
                    //string toDate = d2.AddMinutes((double)(-150)).ToString("yyyy-MM-dd HHmm");
                    string filename = sub.Replace(":", " ") + ".xls";// Convert.ToDateTime(fdate).ToString("yyyy-MM-dd HHmm") + ".xls";
                    string filePath = FileDirectory + @"\" + filename;
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }
                    //傳入模板名稱及sheet名稱
                    //FUSE.ExportExcel.TableToExcel(dt, @"D:\MFROutPutTemplate.xls");
                    MemoryStream ms = FUSE.ExportExcel.GeneratePublicSend(dt, strTemplateName, strSheetName);
                    StreamWriter sw = new StreamWriter(filePath, true, System.Text.Encoding.GetEncoding("GB2312"));
                    ms.WriteTo(sw.BaseStream);
                    sw.Close();
                    ml.Attachment.Add(filePath);
                }
                return true;
            }
            catch (Exception ex)
            {
                WriteLog("Error accurs on method: AddMailBody, error message: " + ex.Message, " CollectInstockOutPutData");
                return false;
            }
        }
        private bool AddMFRMailAttachment3(Mail ml, DataTable dt, string strTemplateName, string strSheetName, string sub)
        {
            try
            {
                string FileDirectory = serviceInstallPath + @"files\InstockHourly\";
                if (!Directory.Exists(FileDirectory))
                {
                    Directory.CreateDirectory(FileDirectory);
                }
                if (dt.Rows.Count > 0)
                {
                    //DateTime d;
                    //DateTime.TryParseExact(fdate, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, out d);
                    //DateTime d2;
                    //DateTime.TryParseExact(tdate, "yyyyMMddHHmmss", CultureInfo.CurrentCulture, DateTimeStyles.None, out d2);

                    //string fromDate = d.AddMinutes((double)(-150)).ToString("yyyy-MM-dd HHmm");
                    //string toDate = d2.AddMinutes((double)(-150)).ToString("yyyy-MM-dd HHmm");
                    string filename = sub + ".xls";// Convert.ToDateTime(fdate).ToString("yyyy-MM-dd HHmm") + ".xls";
                    string filePath = FileDirectory + @"\" + filename;
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }
                    //傳入模板名稱及sheet名稱
                    //FUSE.ExportExcel.TableToExcel(dt, @"D:\MFROutPutTemplate.xls");
                    MemoryStream ms = FUSE.ExportExcel.GeneratePublicSend(dt, strTemplateName, strSheetName);
                    StreamWriter sw = new StreamWriter(filePath, true, System.Text.Encoding.GetEncoding("GB2312"));
                    ms.WriteTo(sw.BaseStream);
                    sw.Close();
                    ml.Attachment.Add(filePath);
                }
                return true;
            }
            catch (Exception ex)
            {
                WriteLog("Error accurs on method: AddMailBody, error message: " + ex.Message, " CollectInstockOutPutData");
                return false;
            }
        }

        public static MemoryStream ARReport(Dictionary<string, string> filter_var, System.Data.DataTable detail_data)
        {
            HSSFWorkbook hssfworkbook = InitializeWorkbook("ARReport.xls", "", filter_var);

            ISheet sheet = hssfworkbook.GetSheet("ARReport");

            int f_row = 0;//Line positioning for passing parameters
            int f_col = 0;//Column targeting for filtering parameters


            foreach (string f_str in filter_var.Keys)
            {
                Search(sheet, string.Format("<{0}>", f_str), true, true, 0, 0, sheet.LastRowNum, 20, ref f_row, ref f_col);
                sheet.GetRow(f_row).GetCell(f_col).SetCellValue(filter_var[f_str].ToString());
            }

            int d_row = 0;//Row positioning for detailed data
            int d_col = 0;//Row positioning for detailed data


            //Total number of rows of detail data

            int d_row_count = detail_data.Rows.Count;

            //Position the start position of data filling

            Search(sheet, "End.[AR Report]", true, true, 0, 0, sheet.LastRowNum, 1, ref d_row, ref d_col);

            //Insert the specified number of blank rows

            InsertRow(sheet, d_row, d_row_count - 1, sheet.GetRow(d_row - 1));

            //The total number of rows of detailed data

            int d_col_count = sheet.GetRow(d_row - 1).LastCellNum;

            #region "Match the position of the source data column and the data column in the template"

            string[] cols = new string[d_col_count];

            for (int index = 0; index < d_col_count; index++)
            {
                cols[index] = sheet.GetRow(d_row - 1).GetCell(index).StringCellValue.Trim("[]".ToCharArray());
            }

            #endregion


            #region "Data input"

            for (int index = 0; index < detail_data.Rows.Count; index++)
            {

                for (int i = 0; i < d_col_count; i++)
                {
                    if (detail_data.Columns.Contains(cols[i]))
                    {
                        switch (detail_data.Columns[cols[i]].DataType.Name)
                        {
                            case "String":
                                sheet.GetRow(d_row - 1 + index).GetCell(i).SetCellValue(detail_data.Rows[index][cols[i]].ToString());
                                break;
                            case "Int":
                            case "Double":
                            case "Decimal":
                            case "Float":
                                if (!Convert.IsDBNull(detail_data.Rows[index][cols[i]]))
                                    sheet.GetRow(d_row - 1 + index).GetCell(i).SetCellValue(System.Convert.ToDouble(detail_data.Rows[index][cols[i]]));
                                else
                                    sheet.GetRow(d_row - 1 + index).GetCell(i).SetCellValue("");
                                break;
                            default:
                                sheet.GetRow(d_row - 1 + index).GetCell(i).SetCellValue(detail_data.Rows[index][cols[i]].ToString());
                                break;

                        }
                    }
                    else
                    {
                        sheet.GetRow(d_row - 1 + index).GetCell(i).SetCellValue("");
                    }
                }
            }

            #endregion

            bool isFill = false;
            ICellStyle style = sheet.GetRow(d_row - 2).GetCell(0).CellStyle;

            for (int index = d_row - 1; index < d_row - 1 + d_row_count; index++)
            {
                isFill = false;

                for (int j = 0; j < sheet.GetRow(index).LastCellNum; j++)
                {

                    if (!isFill && sheet.GetRow(index).GetCell(j).ToString().StartsWith("All"))
                    {
                        isFill = true;
                    }

                    if (isFill)
                    {
                        sheet.GetRow(index).GetCell(j).CellStyle = style;
                    }
                }
            }
            //Write data to memory

            return WriteToStream(hssfworkbook);
        }
        public static MemoryStream WriteToStream(HSSFWorkbook hssfworkbook)
        {
            //Write the stream data of workbook to the root directory
            MemoryStream file = new MemoryStream();


            hssfworkbook.Write(file);
            return file;
        }
        public static void Search(ISheet sheet, string targString, bool excautMatch, bool distinguishUperAndLower, int startRowIndex, int startColIndex, int endRowIndex, int endColIndex, ref int foundRowIndex, ref int foundColIndex)
        {
            try
            {
                if (((startRowIndex > -1) && (startColIndex > -1)) && ((endRowIndex > -1) && (endColIndex > -1)))
                {
                    ISheet sheet2 = sheet;
                    int num = ((startRowIndex - endRowIndex) > 0) ? endRowIndex : startRowIndex;
                    int num2 = ((startColIndex - endColIndex) > 0) ? endColIndex : startColIndex;
                    int num3 = ((startRowIndex - endRowIndex) > 0) ? startRowIndex : (((startRowIndex - endRowIndex) < 0) ? endRowIndex : 0);
                    int num4 = ((startColIndex - endColIndex) > 0) ? startColIndex : (((startColIndex - endColIndex) < 0) ? endColIndex : 0);
                    for (int i = num; i <= (num + num3); i++)
                    {
                        for (int j = num2; j <= (num2 + num4); j++)
                        {
                            ICell cell = null;
                            try
                            {
                                cell = sheet2.GetRow(i).GetCell(j);
                            }
                            catch
                            {
                                continue;
                            }
                            //if (((cell != null) && (cell.CellFormula == "")) && (cell.StringCellValue == targString))
                            if (((cell != null) && (cell.CellType == CellType.STRING)) && (cell.StringCellValue == targString))
                            {

                                foundRowIndex = i;
                                foundColIndex = j;
                                return;
                            }
                        }
                    }
                    foundRowIndex = -1;
                    foundColIndex = -1;
                }
            }
            catch (Exception exception)
            {
                throw new Exception("Internal error\r\n" + exception.ToString(), exception);
            }
        }
        public static void InsertRow(ISheet sheet, int insertRowIndex, int rowQty, IRow SoruceFormateRow)
        {
            sheet.ShiftRows(insertRowIndex, sheet.LastRowNum, rowQty, true, false);
            for (int i = insertRowIndex; i < ((insertRowIndex + rowQty) - 1); i++)
            {
                IRow row = null;
                ICell cell = null;
                ICell cell2 = null;
                row = sheet.CreateRow(i + 1);
                for (int k = SoruceFormateRow.FirstCellNum; k < SoruceFormateRow.LastCellNum; k++)
                {
                    cell = SoruceFormateRow.GetCell(k);
                    if (cell != null)
                    {
                        cell2 = row.CreateCell(k);
                        cell2.CellStyle = cell.CellStyle;
                        //cell2.set_CellStyle(cell.get_CellStyle());
                        cell2.SetCellType(cell.CellType);
                        //cell2.SetCellType(cell.get_CellType());
                    }
                }
            }
            IRow row2 = sheet.GetRow(insertRowIndex);
            ICell cell3 = null;
            ICell cell4 = null;
            for (int j = SoruceFormateRow.FirstCellNum; j < SoruceFormateRow.LastCellNum; j++)
            {
                cell3 = SoruceFormateRow.GetCell(j);
                if (cell3 != null)
                {
                    cell4 = row2.CreateCell(j);
                    cell4.CellStyle = cell3.CellStyle;
                    //cell4.set_CellStyle(cell3.get_CellStyle());
                    cell4.SetCellType(cell3.CellType);
                }
            }
        }
        public static HSSFWorkbook InitializeWorkbook(string template, string subject, Dictionary<string, string> filter_val)
        {
            FileStream file = new FileStream(System.Web.HttpContext.Current.Server.MapPath("../template/") + template, FileMode.Open, FileAccess.Read);


            HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);

            //create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "Foxconn";
            hssfworkbook.DocumentSummaryInformation = dsi;

            //create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = subject;
            si.Author = "Tom.K.Wan";
            si.CreateDateTime = System.DateTime.Now;
            si.Comments = "";
            hssfworkbook.SummaryInformation = si;
            file.Close();
            file.Dispose();
            return hssfworkbook;
        }
        private static void WriteLog(string message, string methodName)
        {
            string logDirectory = serviceInstallPath + "\\log\\" + methodName;
            string strPath = logDirectory + "\\" + DateTime.Today.ToString("yyyy-MM-dd") + ".txt";
            try
            {
                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }
                if (!File.Exists(strPath))
                {
                    FileStream fs = File.Create(strPath);
                    fs.Close();
                }
                StreamWriter sw = File.AppendText(strPath);
                sw.WriteLine(DateTime.Now.ToString() + "\t" + message);
                sw.Flush();
                sw.Close();
                sw.Dispose();
            }
            catch (Exception ex)
            {

            }
        }
        private void DebugStatements(string message)
        {

            string filePath = AppDomain.CurrentDomain.BaseDirectory + "\\Log_" + string.Format("{0:yyyyMMdd}", DateTime.Now) + ".txt";
            lock (typeof(File))
            {
                File.AppendAllText(filePath, DateTime.Now.ToString() + "     " + message + Environment.NewLine);
            }
        }

        public DataSet GetData()
        {
            DataSet myDs = new DataSet();
            using (OracleConnection orcn = new OracleConnection(_DBConnStr))
            {
                OracleCommand cmd = new OracleCommand("yield_report.get_mfr_output_data", orcn);
                OracleDataAdapter oda = new OracleDataAdapter(cmd);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("p_time_start", OracleDbType.Varchar2, 50).Value = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 08:00");
                cmd.Parameters.Add("p_time_end", OracleDbType.Varchar2, 50).Value = DateTime.Now.ToString("yyyy-MM-dd 08:00");
                cmd.Parameters.Add("o_item_data", OracleDbType.RefCursor).Direction = ParameterDirection.Output;
                try
                {
                    orcn.Open();
                    oda.Fill(myDs);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetMFROutPutData, error message: " + ex.Message, "GetMFROutPutData");
                }
            }
            return myDs;
        }

        //Send public SMS
        private void SendPublicSMS2()
        {
            string msgmobiles = string.Empty;
            //Thread.Sleep(15000);
            while (this._runFlag)
            {
                try
                {
                    DataTable dtFamilies = GetOBAFamily("").Tables[0];
                    DataTable mobiles = GetUnSendSMSRecord(1).Tables[0];
                    if (mobiles.Rows.Count > 0)
                    {
                        DataRow row = mobiles.Rows[0];
                        msgmobiles = row["SEND_TO"].ToString();
                    }
                    if (dtFamilies.Rows.Count > 0)
                    {
                        for (int j = 0; j < dtFamilies.Rows.Count; j++)
                        {
                            string message = string.Empty;
                            string family = dtFamilies.Rows[j][0].ToString();
                            string schema = dtFamilies.Rows[j][1].ToString();
                            DataSet unSendSMSRecord = GetTimeGapData(schema, family, 2); //this.GetUnSendSMSRecord();
                            bool flag = unSendSMSRecord.Tables.Count > 0;
                            if (flag)
                            {
                                DataTable dataTable = unSendSMSRecord.Tables[0];
                                for (int i = 0; i < dataTable.Rows.Count; i++)
                                {
                                    DataRow dataRow = dataTable.Rows[i];
                                    message += dataRow["current_station"].ToString() + "=" + dataRow["Qty"].ToString() + ", ";
                                }
                                message = "Family=" + family + ", More than 2hrs not processed packing Level: " + message;
                                string[] array = msgmobiles.TrimEnd(new char[] { ';' }).Split(new char[] { ';' });
                                for (int k = 0; k < array.Length; k++)
                                {
                                    this.sendsms(array[k], message);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    WriteLog("SendPublicSMS, error message: " + ex.Message, "SMS");
                }
                Thread.Sleep(60000);
            }
        }
        private void SendPublicSMS(int timegap, int fseq, int mgroupid)
        {
            string msgmobiles = string.Empty;
            string txtmsg = string.Empty;

            if (timegap == 2)
                txtmsg = "> 2hrs not processed(KEY to MPK): ";
            else if (timegap == 8)
                txtmsg = "> 8hrs not processed(KEY to MPK): ";
            else
                txtmsg = "> 24hrs not processed(KEY to MPK): ";
            try
            {
                DataTable dtFamilies = GetOBAFamily2(fseq).Tables[0];
                DataTable mobiles = GetUnSendSMSRecord(mgroupid).Tables[0];
                if (mobiles.Rows.Count > 0)
                {
                    DataRow row = mobiles.Rows[0];
                    msgmobiles = row["SEND_TO"].ToString();
                }
                if (dtFamilies.Rows.Count > 0)
                {
                    for (int j = 0; j < dtFamilies.Rows.Count; j++)
                    {
                        string message = string.Empty;
                        string family = dtFamilies.Rows[j][0].ToString();
                        string schema = dtFamilies.Rows[j][1].ToString();
                        DataSet unSendSMSRecord = GetTimeGapData(schema, family, timegap); //this.GetUnSendSMSRecord();
                        bool flag = unSendSMSRecord.Tables[0].Rows.Count > 0;
                        if (flag)
                        {
                            DataTable dataTable = unSendSMSRecord.Tables[0];
                            for (int i = 0; i < dataTable.Rows.Count; i++)
                            {
                                DataRow dataRow = dataTable.Rows[i];
                                message += dataRow["current_station"].ToString() + "=" + dataRow["Qty"].ToString() + ", ";
                            }
                            message = family + txtmsg + message;
                            string[] array = msgmobiles.TrimEnd(new char[] { ';' }).Split(new char[] { ';' });
                            for (int k = 0; k < array.Length; k++)
                            {
                                this.sendsms(array[k], message);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog("SendPublicSMS, error message: " + ex.Message, "SMS");
            }
        }
        public void sendsms(string mobno, string msg)
        {
            ServicePointManager.Expect100Continue = false;
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create("http://pay4sms.in/sendsms/?");
            ASCIIEncoding aSCIIEncoding = new ASCIIEncoding();
            string text = string.Empty;
            text = string.Concat(new string[]
			{
				text,
				"token=0faf8b7d9ce712f67d768a7de809c543&credit=2&sender=RSMIPL&message=",
				msg,
				"&number=",
				mobno
			});
            byte[] bytes = aSCIIEncoding.GetBytes(text);
            httpWebRequest.Method = "POST";
            httpWebRequest.ContentType = "application/x-www-form-urlencoded";
            httpWebRequest.ContentLength = (long)bytes.Length;
            using (Stream requestStream = httpWebRequest.GetRequestStream())
            {
                requestStream.Write(bytes, 0, bytes.Length);
            }
            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            string text2 = new StreamReader(httpWebResponse.GetResponseStream()).ReadToEnd();
            Thread.Sleep(1000);
        }
        public DataSet GetTimeGapData(string schemaa, string family, int gaptime)
        {
            DataSet myDs = new DataSet();
            using (OracleConnection orcn = new OracleConnection(_DBConnStr))
            {
                //OracleCommand cmd = new OracleCommand(schemaa + ".yield1.job_collect_timegap_data", orcn);
                OracleCommand cmd = new OracleCommand(schemaa + ".JOB_COLLECT_TIMEGAP_DATA", orcn);
                OracleDataAdapter oda = new OracleDataAdapter(cmd);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("p_family_id", OracleDbType.Varchar2, 50).Value = family;
                cmd.Parameters.Add("p_time", OracleDbType.Varchar2, 50).Value = gaptime;
                cmd.Parameters.Add("o_item_data", OracleDbType.RefCursor).Direction = ParameterDirection.Output;
                try
                {
                    orcn.Open();
                    oda.Fill(myDs);
                }
                catch (Exception ex)
                {
                    WriteLog("Error accurs on method: GetTimeGapData, error message: " + ex.Message, "SMS");
                }
            }
            return myDs;
        }
        private DataSet GetUnSendSMSRecord(int groupid)
        {
            DataSet dataSet = new DataSet();
            using (OracleConnection oracleConnection = new OracleConnection(this._DBConnStr))
            {
                string cmdText = " SELECT * FROM LINE_MONITOR.AUTO_SMS_GROUP A WHERE A.GROUP_ID=" + groupid;
                OracleCommand selectCommand = new OracleCommand(cmdText, oracleConnection);
                try
                {
                    oracleConnection.Open();
                    OracleDataAdapter oracleDataAdapter = new OracleDataAdapter(selectCommand);
                    oracleDataAdapter.Fill(dataSet);
                }
                catch (Exception ex)
                {
                    WriteLog("GetUnSendSMSRecord, error message: " + ex.Message, "SMS");
                }
            }
            return dataSet;
        }

        private void Collect2HoursTimeGapData()
        {
            WriteLog("Start " + DateTime.Now.ToShortDateString(), "Collect2HoursTimeGapData");
            while (_runFlag)
            {
                try
                {
                    int iTaskId = 1011;//Set IND time in DB like 6=>8.30
                    DataTable scheduledTask = this.GetScheduledTask(iTaskId);
                    WriteLog("Called Collect2HoursTimeGapData " + DateTime.Now.ToShortDateString(), "Collect2HoursTimeGapData");
                    bool flag = scheduledTask == null;
                    if (flag)
                    {
                        WriteLog("Wait Schedule " + DateTime.Now.ToShortDateString(), "Schedule");
                        Thread.Sleep(20000);
                        continue;
                    }
                    else
                    {
                        bool flag2 = scheduledTask.Rows.Count == 0;
                        if (flag2)
                        {
                            WriteLog("Wait Schedule " + DateTime.Now.ToShortDateString(), "Schedule");
                            Thread.Sleep(20000);
                            continue;
                        }
                        else
                        {
                            DateTime preDate = Convert.ToDateTime(scheduledTask.Rows[0]["NEXT_TIME"]);
                            int num = Convert.ToInt32(scheduledTask.Rows[0]["TIME_LENGTH"]);
                            //bool flag3 = DateTime.Now.AddMinutes(-10) > preDate.AddMinutes((double)(num));
                            this.MailTriTime = double.Parse(System.Configuration.ConfigurationManager.AppSettings["MailTriggerTimeGap"].ToString());
                            bool flag3 = DateTime.Now.AddMinutes(-MailTriTime) > preDate;
                            if (flag3)
                            {

                                try
                                {
                                    DateTime scheduledTime = preDate.AddMinutes((double)num);
                                    WriteLog("If start " + DateTime.Now.ToShortDateString(), "UpdateSchedualedTaskIF");
                                    SendPublicSMS(2, 2, 6);
                                    this.UpdateSchedualedTask(iTaskId, scheduledTime, "OK", "");
                                    WriteLog("If end " + DateTime.Now.ToShortDateString(), "UpdateSchedualedTaskIF");
                                }
                                catch (Exception ex2)
                                {
                                    WriteLog("Collect2HoursTimeGapData " + ex2.ToString() + " " + DateTime.Now.ToString(), "Exception");
                                }
                            }
                        }
                    }

                }
                catch (Exception ex4)
                {
                    WriteLog(ex4.ToString() + DateTime.Now.ToShortDateString(), "Exception");
                }
                Thread.Sleep(20000);

            }
            WriteLog("End", "Collect2HoursTimeGapData...." + DateTime.Now.ToShortDateString());
        }
        private void Collect8HoursTimeGapData()
        {
            WriteLog("Start " + DateTime.Now.ToShortDateString(), "Collect8HoursTimeGapData");
            while (_runFlag)
            {
                try
                {
                    int iTaskId = 1012;//Set IND time in DB like 6=>8.30
                    DataTable scheduledTask = this.GetScheduledTask(iTaskId);
                    WriteLog("Called Collect8HoursTimeGapData " + DateTime.Now.ToShortDateString(), "Collect8HoursTimeGapData");
                    bool flag = scheduledTask == null;
                    if (flag)
                    {
                        WriteLog("Wait Schedule " + DateTime.Now.ToShortDateString(), "Schedule");
                        Thread.Sleep(20000);
                        continue;
                    }
                    else
                    {
                        bool flag2 = scheduledTask.Rows.Count == 0;
                        if (flag2)
                        {
                            WriteLog("Wait Schedule " + DateTime.Now.ToShortDateString(), "Schedule");
                            Thread.Sleep(20000);
                            continue;
                        }
                        else
                        {
                            DateTime preDate = Convert.ToDateTime(scheduledTask.Rows[0]["NEXT_TIME"]);
                            int num = Convert.ToInt32(scheduledTask.Rows[0]["TIME_LENGTH"]);
                            //bool flag3 = DateTime.Now.AddMinutes(-10) > preDate.AddMinutes((double)(num));
                            this.MailTriTime = double.Parse(System.Configuration.ConfigurationManager.AppSettings["MailTriggerTimeGap"].ToString());
                            bool flag3 = DateTime.Now.AddMinutes(-MailTriTime) > preDate;
                            if (flag3)
                            {

                                try
                                {
                                    DateTime scheduledTime = preDate.AddMinutes((double)num);
                                    WriteLog("If start " + DateTime.Now.ToShortDateString(), "UpdateSchedualedTaskIF");
                                    SendPublicSMS(8, 2, 7);
                                    this.UpdateSchedualedTask(iTaskId, scheduledTime, "OK", "");
                                    WriteLog("If end " + DateTime.Now.ToShortDateString(), "UpdateSchedualedTaskIF");
                                }
                                catch (Exception ex2)
                                {
                                    WriteLog("Collect8HoursTimeGapData " + ex2.ToString() + " " + DateTime.Now.ToString(), "Exception");
                                }
                            }
                        }
                    }

                }
                catch (Exception ex4)
                {
                    WriteLog(ex4.ToString() + DateTime.Now.ToShortDateString(), "Exception");
                }
                Thread.Sleep(20000);

            }
            WriteLog("End", "Collect8HoursTimeGapData...." + DateTime.Now.ToShortDateString());
        }
        private void Collect24HoursTimeGapData()
        {
            WriteLog("Start " + DateTime.Now.ToShortDateString(), "Collect24HoursTimeGapData");
            while (_runFlag)
            {
                try
                {
                    int iTaskId = 1013;//Set IND time in DB like 6=>8.30
                    DataTable scheduledTask = this.GetScheduledTask(iTaskId);
                    WriteLog("Called GetScheduledTask Method " + DateTime.Now.ToShortDateString(), "Collect24HoursTimeGapData");
                    bool flag = scheduledTask == null;
                    if (flag)
                    {
                        WriteLog("Wait Schedule " + DateTime.Now.ToShortDateString(), "Schedule");
                        Thread.Sleep(20000);
                        continue;
                    }
                    else
                    {
                        bool flag2 = scheduledTask.Rows.Count == 0;
                        if (flag2)
                        {
                            WriteLog("Wait Schedule " + DateTime.Now.ToShortDateString(), "Schedule");
                            Thread.Sleep(20000);
                            continue;
                        }
                        else
                        {
                            DateTime preDate = Convert.ToDateTime(scheduledTask.Rows[0]["NEXT_TIME"]);
                            int num = Convert.ToInt32(scheduledTask.Rows[0]["TIME_LENGTH"]);
                            //bool flag3 = DateTime.Now.AddMinutes(-10) > preDate.AddMinutes((double)(num));
                            this.MailTriTime = double.Parse(System.Configuration.ConfigurationManager.AppSettings["MailTriggerTimeGap"].ToString());
                            bool flag3 = DateTime.Now.AddMinutes(-MailTriTime) > preDate;
                            if (flag3)
                            {

                                try
                                {
                                    DateTime scheduledTime = preDate.AddMinutes((double)num);
                                    WriteLog("If start " + DateTime.Now.ToShortDateString(), "UpdateSchedualedTaskIF");
                                    SendPublicSMS(24, 2, 8);
                                    this.UpdateSchedualedTask(iTaskId, scheduledTime, "OK", "");
                                    WriteLog("If end " + DateTime.Now.ToShortDateString(), "UpdateSchedualedTaskIF");
                                }
                                catch (Exception ex2)
                                {
                                    WriteLog("Collect24HoursTimeGapData " + ex2.ToString() + " " + DateTime.Now.ToString(), "Exception");
                                }
                            }
                        }
                    }

                }
                catch (Exception ex4)
                {
                    WriteLog(ex4.ToString() + DateTime.Now.ToShortDateString(), "Exception");
                }
                Thread.Sleep(20000);
            }
            WriteLog("End Collect24HoursTimeGapData...." + DateTime.Now.ToShortDateString(), "Collect24HoursTimeGapData");
        }

    }
}
