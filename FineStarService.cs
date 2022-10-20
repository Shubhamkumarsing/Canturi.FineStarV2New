using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Threading;
using Timer = System.Threading.Timer;



namespace Canturi.FineStarV2
{
    public partial class FineStarService : ServiceBase
    {
        string _constring = string.Empty;
        

        public FineStarService()
        {
            InitializeComponent();
             
            
            try
            {    //Pass the connection String for the Service
                _constring = ConfigurationManager.AppSettings["CanturiConnectionStr"].ToString();
            }
            catch
            {
                LogError("Sorry, connection to database could not made");

            }
           
           //FineStarDiamond();

        }


       
        //To Write into the Log fife
        public static void LogError(string msg)
        {
            StreamWriter str = new StreamWriter(ConfigurationManager.AppSettings["LogFilePath"].ToString(), true);
            if (msg == "") {
                str.WriteLine(DateTime.Now.ToString());
            }
            else
            {
                str.WriteLine(msg);
            }

            
            str.Close();
            str.Dispose();
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                LogError("FineStar Service started "+ DateTime.Now.ToString());
                this.ServiceTimer_Tick();

            }
            catch (Exception ex)
            {
                LogError("Error Timer On Start - " + DateTime.Now.ToString() + " - " + ex.Message.ToString() + " - " + ex.StackTrace.ToString());
            }
        }

        protected override void OnStop()
        
          {
           LogError("FineStar Service stopped {0}");
            this.Schedular.Dispose();
        }


        //Time Schdscheduler to Start the Service Ones in a days
        private Timer Schedular;
        private void ServiceTimer_Tick()
        {
            try
            {
                Schedular = new Timer(new TimerCallback(SchedularCallback));
                string mode = ConfigurationManager.AppSettings["Mode"].ToUpper();
                LogError("FineStar Service Mode: " + mode + "");

                //Set the Default Time.
                DateTime scheduledTime = DateTime.MinValue;

                if (mode == "DAILY")
                {
                    //Get the Scheduled Time from AppSettings.
                    scheduledTime = DateTime.Parse(System.Configuration.ConfigurationManager.AppSettings["ScheduledTime"]);
                    if (DateTime.Now > scheduledTime)
                    {
                        //If Scheduled Time is passed set Schedule for the next day.
                        scheduledTime = scheduledTime.AddDays(1);
                     FineStarDiamond();
                    }
                }

               // if (mode.ToUpper() == "INTERVAL")
              //  {
                    //Get the Interval in Minutes from AppSettings.
                  //  int intervalMinutes = Convert.ToInt32(ConfigurationManager.AppSettings["IntervalMinutes"]);

                    //Set the Scheduled Time by adding the Interval to Current Time.
                  //  scheduledTime = DateTime.Now.AddMinutes(intervalMinutes);
                  //  if (DateTime.Now > scheduledTime)
                  //  {
                        //If Scheduled Time is passed set Schedule for the next Interval.
                      //  scheduledTime = scheduledTime.AddMinutes(intervalMinutes);
                   // }
               // }

                TimeSpan timeSpan = scheduledTime.Subtract(DateTime.Now);
                string schedule = string.Format("{0} Days,{1} Hours,{2} minutes,{3} Seconds", timeSpan.Days, timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);

               LogError("FineStar Service scheduled to next run after: " + schedule + " {0}");
                LogError("WebService Session end    " + DateTime.Now.ToString());

                //Get the difference in Minutes between the Scheduled and Current Time.
                int dueTime = Convert.ToInt32(timeSpan.TotalMilliseconds);

                //Change the Timer's Due Time.
                Schedular.Change(dueTime, Timeout.Infinite);
            }
            catch (Exception ex)
                {
                    LogError("Error Timer Handler - " + DateTime.Now.ToString() + " - " + ex.Message.ToString() + " - " + ex.Source.ToString());
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("SimpleService"))
                {
                    serviceController.Stop();
                }
            }
               
        }
        //To schedule the Next Runtime Of Service
        private void SchedularCallback(object e)
        {
            LogError("FineStar Service Log:"+DateTime.Now.ToString());
            this.ServiceTimer_Tick();
        }

       

        // For Diamond Validation 
        public bool IsValidDiamond(string luster, string location, string clarity, string status)
        {
            if (location.ToUpper() == "INDIA")
            {
                if (string.IsNullOrEmpty(clarity))
                {
                    clarity = "";
                }

                string[] invalidClarity = { "SI2", "SI1", "VS2" };
                if (invalidClarity.Contains(clarity.Trim().ToUpper()))
                {
                    return false;
                }
            }
            if (!String.IsNullOrEmpty(status))
            {
                if (status.ToUpper() == "B")
                {
                    return false;
                }
            }

            if (String.IsNullOrEmpty(luster))
            {
                return true;
            }
            else
            {
                if (luster.ToUpper() == "M0")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }


        }


        public int AddUpFineStarDiamond(RootObject objDiamondModels, int rowNumber)
        {
            int result = -1;

            try
            {

                SqlParameter prmREPORT_NO = SqlHelper.CreateParameter("@REPORT_NO", objDiamondModels.REPORTNO);
                SqlParameter prmCERTI_IMAGE = SqlHelper.CreateParameter("@CERTI_IMAGE", objDiamondModels.CERTIIMAGE);
                SqlParameter prmHEART_IMAGE = SqlHelper.CreateParameter("@HEART_IMAGE", objDiamondModels.HEARTIMAGE);
                SqlParameter prmPLOTTING_IMAGE = SqlHelper.CreateParameter("@PLOTTING_IMAGE", objDiamondModels.PLOTTINGIMAGE); ;
                SqlParameter prmPURITY = SqlHelper.CreateParameter("@PURITY", objDiamondModels.PURITY);
                SqlParameter prmCOLOR = SqlHelper.CreateParameter("@COLOR", objDiamondModels.COLOR);
                SqlParameter prmREPORT_COMMENT = SqlHelper.CreateParameter("@REPORT_COMMENT", objDiamondModels.REPORTCOMMENT);
                
                SqlParameter prmCROWN_ANGLE = SqlHelper.CreateParameter("@CROWN_ANGLE", objDiamondModels.CROWNANGLE);
                SqlParameter prmCROWN_HEIGHT = SqlHelper.CreateParameter("@CROWN_HEIGHT", objDiamondModels.CROWNHEIGHT);
                SqlParameter prmCULET = SqlHelper.CreateParameter("@CULET", objDiamondModels.CULET);
                SqlParameter prmCUT = SqlHelper.CreateParameter("@CUT", GetCut(objDiamondModels.CUT));
                SqlParameter prmDEPTH_PER = SqlHelper.CreateParameter("@DEPTH_PER", objDiamondModels.DEPTHPER);
                SqlParameter prmREAL_IMAGE = SqlHelper.CreateParameter("@REAL_IMAGE", objDiamondModels.REALIMAGE);
                SqlParameter prmEYE_CLEAN = SqlHelper.CreateParameter("@EYE_CLEAN", GetEyeClean(objDiamondModels.EYECLEAN));

                SqlParameter prmFLS = SqlHelper.CreateParameter("@FLS", GetFluorescence(objDiamondModels.FLS));
                SqlParameter prmKEY_TO_SYMBOLS = SqlHelper.CreateParameter("@KEY_TO_SYMBOLS", objDiamondModels.KEYTOSYMBOLS);
                SqlParameter prmLAB = SqlHelper.CreateParameter("@LAB", objDiamondModels.LAB);
                SqlParameter prmLASER_INSCRIPTION = SqlHelper.CreateParameter("@LASER_INSCRIPTION", objDiamondModels.LASERINSCRIPTION);
                SqlParameter prmMEASUREMENT = SqlHelper.CreateParameter("@MEASUREMENT", FnMeasurements(objDiamondModels.LENGTH1, objDiamondModels.WIDTH,objDiamondModels.DEPTH)) ;
                SqlParameter prmLENGTH_1 = SqlHelper.CreateParameter("@LENGTH_1", objDiamondModels.LENGTH1);
                SqlParameter prmWIDTH = SqlHelper.CreateParameter("@WIDTH", objDiamondModels.WIDTH);

                SqlParameter prmDEPTH = SqlHelper.CreateParameter("@DEPTH", objDiamondModels.DEPTH);
                SqlParameter prmPAV_ANGLE = SqlHelper.CreateParameter("@PAV_ANGLE", objDiamondModels.PAVANGLE);
                SqlParameter prmPAV_HEIGHT = SqlHelper.CreateParameter("@PAV_HEIGHT", objDiamondModels.PAVHEIGHT);
                SqlParameter prmPOLISH = SqlHelper.CreateParameter("@POLISH", GetPolish(objDiamondModels.POLISH));
                SqlParameter prmRAP_PRICE = SqlHelper.CreateParameter("@RAP_PRICE", objDiamondModels.RAPPRICE);
                SqlParameter prmCTS = SqlHelper.CreateParameter("@CTS", objDiamondModels.CTS);
                SqlParameter prmAVG_DIA = SqlHelper.CreateParameter("@AVG_DIA", objDiamondModels.AVGDIA);

                SqlParameter prmDISC_PER = SqlHelper.CreateParameter("@DISC_PER", objDiamondModels.DISCPER);
                SqlParameter prmNET_RATE = SqlHelper.CreateParameter("@NET_RATE", objDiamondModels.NETRATE);
                SqlParameter prmNET_VALUE = SqlHelper.CreateParameter("@NET_VALUE", objDiamondModels.NETVALUE);
                SqlParameter prmSHAPE = SqlHelper.CreateParameter("@SHAPE", GetShape(objDiamondModels.SHAPE));
                SqlParameter prmSHADE = SqlHelper.CreateParameter("@SHADE", objDiamondModels.SHADE);
                SqlParameter prmSYMM = SqlHelper.CreateParameter("@SYMM", GetSymmetry(objDiamondModels.SYMM));
                SqlParameter prmTABLE_PER = SqlHelper.CreateParameter("@TABLE_PER", objDiamondModels.TABLEPER);

                SqlParameter prmHA = SqlHelper.CreateParameter("@HA", objDiamondModels.HA);
                SqlParameter prmPACKET_NO = SqlHelper.CreateParameter("@PACKET_NO", objDiamondModels.PACKETNO);
                SqlParameter prmLOCATION = SqlHelper.CreateParameter("@LOCATION", objDiamondModels.LOCATION);
                SqlParameter prmCENTER_NATTS = SqlHelper.CreateParameter("@CENTER_NATTS", objDiamondModels.CENTERNATTS);
                SqlParameter prmSIDE_NATTS = SqlHelper.CreateParameter("@SIDE_NATTS", objDiamondModels.SIDEFEATHER);
                SqlParameter prmCENTER_FEATHER = SqlHelper.CreateParameter("@CENTER_FEATHER", objDiamondModels.CENTERFEATHER);
                SqlParameter prmSIDE_FEATHER = SqlHelper.CreateParameter("@SIDE_FEATHER", objDiamondModels.SIDEFEATHER);

                SqlParameter prmCROWN_OPEN = SqlHelper.CreateParameter("@CROWN_OPEN", objDiamondModels.CROWNOPEN);
                SqlParameter prmPOP = SqlHelper.CreateParameter("@POP", objDiamondModels.POP);
                SqlParameter prmXTOP = SqlHelper.CreateParameter("@XTOP", objDiamondModels.XTOP);
                SqlParameter prmBRAND = SqlHelper.CreateParameter("@BRAND", objDiamondModels.BRAND);
                SqlParameter prmTYPE2_CERT = SqlHelper.CreateParameter("@TYPE2_CERT", objDiamondModels.TYPE2CERT);
                SqlParameter prmBRILLIANCY = SqlHelper.CreateParameter("@BRILLIANCY", objDiamondModels.BRILLIANCY);
                SqlParameter prmDNA = SqlHelper.CreateParameter("@DNA", objDiamondModels.DNA);
                SqlParameter prmARROW_IMAGE = SqlHelper.CreateParameter("@ARROW_IMAGE", objDiamondModels.ARROWIMAGE);
                //SqlParameter prmASST_SCOPE_IMAGE = SqlHelper.CreateParameter("@ASST_SCOPE_IMAGE", objDiamondModels.ASSTSCOPEIMAGE);
                
                
                SqlParameter prmB2C_MP4 = SqlHelper.CreateParameter("@B2C_MP4", objDiamondModels.B2BMP4);
                SqlParameter prmB2C_IMAGE = SqlHelper.CreateParameter("@B2C_IMAGE", objDiamondModels.B2CIMAGE);
                SqlParameter prmGIRDLE = SqlHelper.CreateParameter("@GIRDLE", objDiamondModels.GIRDLE);
                SqlParameter prmCOMMENTS = SqlHelper.CreateParameter("@COMMENTS", objDiamondModels.COMMENTS);
                SqlParameter prmCERTI_LINK = SqlHelper.CreateParameter("@CERTI_LINK", objDiamondModels.CERTILINK);
                SqlParameter prmVIDEO = SqlHelper.CreateParameter("@VIDEO", objDiamondModels.VIDEO);
                SqlParameter prmLotNumber = SqlHelper.CreateParameter("@LotNumber", objDiamondModels.REPORTNO);
              
                
                //SqlParameter prmIsDiamondOrderd = SqlHelper.CreateParameter("@IsDiamondOrderd", GetIsDiamondOrderd(objDiamondModels.IsDiamondOrderd));


                SqlParameter[] allParams = {
                    prmREPORT_NO,prmCERTI_IMAGE,
                    prmHEART_IMAGE,prmPLOTTING_IMAGE,prmPURITY,prmCOLOR,
                    prmREPORT_COMMENT,prmCROWN_ANGLE,prmCROWN_HEIGHT,prmCULET,prmCUT,
                    prmDEPTH_PER,prmREAL_IMAGE,prmEYE_CLEAN,prmFLS,prmKEY_TO_SYMBOLS,
                    prmLAB,prmLASER_INSCRIPTION,prmMEASUREMENT,prmLENGTH_1,prmWIDTH,
                    prmDEPTH,prmPAV_ANGLE,prmPAV_HEIGHT,prmPOLISH,prmRAP_PRICE,prmCTS,
                    prmAVG_DIA,prmDISC_PER,prmNET_RATE,prmNET_VALUE,prmSHAPE,prmSHADE,
                    prmSYMM,prmTABLE_PER,prmHA,prmPACKET_NO,prmLOCATION,prmCENTER_NATTS,
                    prmSIDE_NATTS,prmCENTER_FEATHER,prmSIDE_FEATHER,prmCROWN_OPEN,prmPOP,
                    prmXTOP,prmBRAND,prmTYPE2_CERT,prmBRILLIANCY,prmDNA,prmARROW_IMAGE,
                    //prmASST_SCOPE_IMAGE,
                    prmB2C_MP4,prmB2C_IMAGE,prmGIRDLE,prmCOMMENTS,
                    prmCERTI_LINK,prmVIDEO,prmLotNumber

                };
                SqlHelper.ExecuteNonQuery(_constring, CommandType.StoredProcedure, "Usp_AddUpdFineStarDiamondV2", allParams);



            }
            catch (Exception ex)
            {
                LogError(ex.Message);
            }   

            return result;
        }


        public void FineStarDiamond()
        {
            try
            {

                //TruncateTable Then Insert
                //SqlHelper.Truncatetableolddata();

                string jsonData = "";
                LogError("WebClient start    " + DateTime.Now.ToString());
                string Username = ConfigurationManager.AppSettings["ApiUserName"].ToString();
                string Password = ConfigurationManager.AppSettings["ApiPassword"].ToString();


                WebRequest request = WebRequest.Create(ConfigurationManager.AppSettings["FinestarApiUrl"].ToString() + "?username=" + Username + "&password=" + Password);
                request.Method = "GET";
                WebResponse response = request.GetResponse();

                var encoding = ASCIIEncoding.ASCII;

                using (var reader = new System.IO.StreamReader(response.GetResponseStream(), encoding))
                {
                    jsonData = reader.ReadToEnd();
                }
                if (!String.IsNullOrEmpty(jsonData))
                {
                    //JsonFolderPath
                    string fileName = Guid.NewGuid().ToString() + ".json";
                    System.IO.File.WriteAllText(ConfigurationManager.AppSettings["JsonFolderPath"].ToString() + fileName, jsonData);
                    LogError("Json file save: " + ConfigurationManager.AppSettings["JsonFolderPath"].ToString() + fileName);



                    List<RootObject> rootData = JsonConvert.DeserializeObject<List<RootObject>>(jsonData);

                    int count = 1;
                    foreach (var item in rootData)
                    {

                        AddUpFineStarDiamond(item, count);
                        count++;

                    }
                    LogError("Successfully Inserting Diamond Procedure");
                }
                else
                {
                    LogError("Empty jsonData");
                }

                
            }



            catch (Exception ex)
            {

                LogError(ex.Message);
            }



        }
     

        public string FnMeasurements(string length, string width, string depth)
        {
            return length + " x " + width + " x " + depth;
        }

        public string GetShape(string shape)
        {
            switch (shape)
            {
                case "ROUND":
                    shape = "Round";
                    break;
                case "PC":
                    shape = "Princess";
                    break;
                case "CML":
                    shape = "Cushion";
                    break;
                case "CS":
                    shape = "Cushion";
                    break;
                case "CM":
                    shape = "Cushion";
                    break;
                case "EM":
                    shape = "Emerald";
                    break;
                case "RN":
                    shape = "Radiant";
                    break;
                case "PEAR":
                    shape = "Pear";
                    break;
                case "OV":
                    shape = "Oval";
                    break;
                case "HT":
                    shape = "Heart";
                    break;
                case "SE":
                    shape = "Asscher";
                    break;
                default:
                    shape = shape;
                    break;
            }
            return shape;
        }

        public string GetCut(string cut)
        {
            switch (cut)
            {
                case "G":
                    cut = "GOOD";
                    break;
                case "VG":
                    cut = "VERY GOOD";
                    break;
                case "EX":
                    cut = "EXCELLENT";
                    break;
                default:
                    cut = cut;
                    break;
            }
            return cut;
        }

        public string GetPolish(string polish)
        {
            switch (polish)
            {
                case "G":
                    polish = "GOOD";
                    break;
                case "VG":
                    polish = "VERY GOOD";
                    break;
                case "EX":
                    polish = "EXCELLENT";
                    break;
                default:
                    polish = polish;
                    break;
            }
            return polish;
        }
        public string GetEyeClean(string EyeClean)
        {
            switch (EyeClean)
            {
                case "Y":
                    EyeClean = "YES";
                    break;
                case "N":
                    EyeClean = "NO";
                    break;
               
                default:
                    EyeClean = EyeClean;
                    break;
            }
            return EyeClean;
        }


        public string GetSymmetry(string symmetry)
        {
            switch (symmetry)
            {
                case "G":
                    symmetry = "GOOD";
                    break;
                case "VG":
                    symmetry = "VERY GOOD";
                    break;
                case "EX":
                    symmetry = "EXCELLENT";
                    break;
                default:
                    symmetry = symmetry;
                    break;
            }
            return symmetry;
        }

        public string GetFluorescence(string fluorescence)
        {
            switch (fluorescence)
            {
                case "NON":
                    fluorescence = "NONE";
                    break;
                case "FNT":
                    fluorescence = "FAINT";
                    break;
                case "SLT":
                    fluorescence = "STORNG FAINT";
                    break;
                case "MED":
                    fluorescence = "MEDIUM";
                    break;
                case "STG":
                    fluorescence = "STRONG";
                    break;
                case "VST":
                    fluorescence = "VERY STRONG";
                    break;
                default:
                    fluorescence = fluorescence;
                    break;
            }
            return fluorescence;
        }



        public string FnEyeClean(string clarity, string location)
        {
            string[] hongKongEyeClean = { "S12", "SI1", "VS2", "VS1", "VVS2", "VVS1", "IF", "FL" };
            string[] indiaEyeClean = { "VS1", "VVS2", "VVS1", "IF", "FL" };
            if (location.ToUpper() == "HONG KONG")
            {
                if (hongKongEyeClean.Contains(clarity.ToUpper()))
                {
                    return "YES";
                }
            }
            if (location.ToUpper() == "INDIA")
            {
                if (indiaEyeClean.Contains(clarity.ToUpper()))
                {
                    return "YES";
                }
            }
            return "NO";
        }







    }
}
