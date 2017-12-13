using System;
using System.Collections.Generic;
using System.ServiceModel.Web;
using System.Text;
using System.IO;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Crm.Sdk.Messages;
using System.Configuration;
using System.Net;
using System.Collections.Specialized;
using Newtonsoft.Json;
using System.ServiceModel.Description;

namespace WcfReferralService
{
    public class ReferralService : IReferralService
    {
        #region Global Variables
        private String userName = ConfigurationManager.AppSettings["user"].ToString();
        private String password = ConfigurationManager.AppSettings["pass"].ToString();
        private String domain = ConfigurationManager.AppSettings["domain"].ToString();
        private Int16 SQLServerTimeoutSeconds = Convert.ToInt16(ConfigurationManager.AppSettings["SQLServerTimeoutSeconds"].ToString());
        public String ErrorMessage = String.Empty;
        public StringBuilder myLog = new StringBuilder();
        #endregion

        [WebInvoke(Method = "POST", UriTemplate =
            "ReferralInfo"
            , RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped)]
        public string ReferralInfo(
            Int32 ReferralId,
            Byte Type,
            String ReferralDate,
            String AdvisorName,
            String ContactName,
            String ContactAddress,
            String ContactEmail,
            String ContactPhone,
            String ProspectName
            )
        {

            myLog = new StringBuilder();
            String returnedErrorDetails = String.Empty;
            String returnedRejectDupeResult = String.Empty;
            String returnedAcceptResult = String.Empty;
            string authInfo = "UserName:Password";
            myLog.AppendLine("Execute ReferralInfo web method.");
            try
            {
                myLog.AppendLine("Parameters retrieved: ");
                myLog.AppendLine("ReferralId: " + ReferralId.ToString());
                myLog.AppendLine("Type: " + Type.ToString());
                myLog.AppendLine("ReferralDate: " + ReferralDate);
                myLog.AppendLine("AdvisorName: " + AdvisorName);
                myLog.AppendLine("ContactName: " + ContactName);
                myLog.AppendLine("ContactAddress: " + ContactAddress);
                myLog.AppendLine("ContactEmail: " + ContactEmail);
                myLog.AppendLine("ContactPhone: " + ContactPhone);
                myLog.AppendLine("ProspectName: " + ProspectName);
                try
                {
                    myGetService(ref myLog); //authenticates to crm api
                }
                catch (Exception e)
                {
                    Results refInfo = new Results
                    {
                        Result = myLog.ToString(),
                        ErrorMessage = "Can't Authenticate to CRM. Please verify that the credentials are valid. Message: " + e.Message
                    };

                    string jsonResult = JsonConvert.SerializeObject(refInfo);
                    return jsonResult;
                }

                #region Search for existing referral in crm system
                myLog.AppendLine("Search for existing referral.");
                var fetchXml =
                        @" 
                            <fetch mapping='logical' distinct='true'>
                              <entity name='contact'>
                                <attribute name='contactid' />
                                <attribute name='createdon' />
                                <attribute name='createdby' />
                                <attribute name='lastname' />
                                    <filter type='and'>
                                      <condition attribute='contactid' operator='eq' value='" + ReferralId + @"' />
                                    </filter>                                    
                              </entity>
                            </fetch>";
                var fetchExpression = new FetchExpression(fetchXml);
                EntityCollection retrievedCont = service.RetrieveMultiple(fetchExpression);
                // Convert the FetchXML into a query expression.
                var conversionRequest = new FetchXmlToQueryExpressionRequest
                {
                    FetchXml = fetchXml
                };
                var conversionResponse =
                    (FetchXmlToQueryExpressionResponse)service.Execute(conversionRequest);
                // Use the newly converted query expression to make a retrieve multiple
                // request to Microsoft Dynamics CRM.
                QueryExpression queryExpression = conversionResponse.Query;
                EntityCollection result = service.RetrieveMultiple(queryExpression);
                myLog.AppendLine("result.Entities.Count.: " + result.Entities.Count.ToString());
                #endregion
                if (result.Entities.Count > 0)
                {
                    myLog.AppendLine("Referral already exists");
                    #region Send Reject Dupe to 3rd party Referral service
                    myLog.AppendLine("Send Reject Dupe to 3rd party Referral service.");

                    var httpWebRequest = (HttpWebRequest)WebRequest.Create("https://www.somewebservice.com/api/RejectDupeReferral");
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = "POST";
                    httpWebRequest.Headers.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.ASCII.GetBytes(authInfo)));

                    using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                    {
                        string json = "[{\"ReferralId\":\"" + ReferralId + "\"," +
                                      "\"OrganizationId\":\"" + organizationID + "\"}]";

                        streamWriter.Write(json);
                        streamWriter.Flush();
                        streamWriter.Close();

                        var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                        using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                        {
                            var result = streamReader.ReadToEnd();
                            myLog.AppendLine("RejectDupeReferral result: " + result.ToString());
                            returnedRejectDupeResult = result.ToString();
                        }
                    }
                    #endregion
                }
                else
                {
                    #region Send Accept to 3rd party referral service
                    myLog.AppendLine("Send Accept to 3rd party Referral service.");
                    var httpWebRequest = (HttpWebRequest)WebRequest.Create("https://www.somewebservice.com/api/AcceptReferral");
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = "POST";
                    httpWebRequest.Headers.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.ASCII.GetBytes(authInfo)));
                    using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                    {
                        string json = "[{\"ReferralId\":\"" + ReferralId + "\"," +
                            "\"OrganizationId\":\"" + organizationID + "\"}]";

                        streamWriter.Write(json);
                        streamWriter.Flush();
                        streamWriter.Close();

                        var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                        RootObject firstRootObject = new RootObject();

                        using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                        {
                            var result = streamReader.ReadToEnd();
                            returnedAcceptResult = result.ToString();
                            myLog.AppendLine("Accept result: " + result.ToString());

                            RootObject[] root = JsonConvert.DeserializeObject<RootObject[]>(result);
                            firstRootObject = root[0];

                            List<object> errList = new List<object>();
                            errList = firstRootObject.Errors;
                            foreach (var obj in errList)
                            {
                                returnedErrorDetails += obj.ToString();
                            }
                            if (returnedErrorDetails != String.Empty)
                            {
                                ErrorMessage += "Error received from 3rd party Referral service. ";
                            }
                        }

                        if (returnedErrorDetails == String.Empty)
                        {
                            if (firstRootObject.Referral.Type != 1) //!= new prospect
                            {
                                ErrorMessage += "Expects Type of 1.  Type sent: " + firstRootObject.Referral.Type.ToString();
                            }
                            else
                            {
                                //create records in crm system
                            }
                        }
                    }
                    #endregion
                }

                #region Log any errors for customer
                if (returnedRejectDupeResult != String.Empty || returnedErrorDetails != String.Empty || ErrorMessage != String.Empty)
                {
                    //Log errors for customer
                    Entity note = new Entity(Annotation.EntityLogicalName);
                    if (returnedErrorDetails != String.Empty)
                    {
                        note["subject"] = "3rd Party Accept result Error for ReferralId: " + ReferralId.ToString();
                        note["notetext"] = "Detail Logs: " + myLog.ToString() + "\r\n 3rd Party Accept Error: " + returnedErrorDetails;
                    }
                    else if (returnedRejectDupeResult != String.Empty)
                    {
                        note["subject"] = "3rd Party RejectDuplicate result for ReferralId: " + ReferralId.ToString();
                        note["notetext"] = "Detail Logs: " + myLog.ToString() + "\r\n RejectDuplicate Result: " + returnedRejectDupeResult;
                    }
                    else //error returned from this service/crm
                    {
                        note["subject"] = "Targeted ReferralInfo Error for ReferralId: " + ReferralId.ToString();
                        note["notetext"] = "Detail Logs: " + myLog.ToString() + "\r\n Error: " + Error;
                    }
                    service.Create(note);
                    myLog.AppendLine("Note log created in CRM.");
                }
                #endregion

                //Send errors returned from this service/crm to the 3rd Party Web Service
                Results refInfoResults = new Results
                {
                    Result = myLog.ToString(),
                    Error = ErrorMessage
                };
                string json4 = JsonConvert.SerializeObject(refInfoResults);
                return json4;
            }
            catch (Exception e)
            {
                Results refInfoResults = new Results
                {
                    Result = myLog.ToString(),
                    Error = "Message: " + e.Message + "Error: " + ErrorMessage
                };

                //Log service errors for customer
                Entity note = new Entity(Annotation.EntityLogicalName);
                note["subject"] = "Targeted ReferralInfo Error for ReferralId: " + ReferralId.ToString();
                note["notetext"] = "Detail Logs: " + myLog.ToString() + "\r\n Message: " + e.Message + "\r\n Error: " + ErrorMessage + "returnedAcceptResult: " + returnedAcceptResult;
                service.Create(note);

                string jsonResult = JsonConvert.SerializeObject(refInfoResults);
                return jsonResult;

            }
        }

        public NameValueCollection TypeCollection()
        {
            NameValueCollection objType = new NameValueCollection();
            objType.Add("Prospect", "1");
            objType.Add("Updated Prospect", "2");
            objType.Add("Deactivated Prospect", "3");
            objType.Add("Prospect Moved", "4");

            return objType;
        }

        protected void myGetService(ref StringBuilder myLog)
        {
            //connect to crm api
        }


        public class Referral
        {
            public int ReferralId { get; set; }
            public byte Type { get; set; }
            public string ReferralDate { get; set; }
            public string AdvisorName { get; set; }
            public string ContactName { get; set; }
            public string ContactAddress { get; set; }
            public string ContactEmail { get; set; }
            public string ContactPhone { get; set; }
            public string ProspectName { get; set; }
        }


        public class RootObject
        {
            [JsonProperty("Errors")]
            public List<object> Errors { get; set; }

            [JsonProperty("ReferralId")]
            public int ReferralId { get; set; }

            [JsonProperty("Referral")]
            public Referral Referral { get; set; }
        }

        public class Results
        {
            public string Result { get; set; }
            public string Error { get; set; }
        }
    }
}

