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

namespace WcfLeadService
{
    public class LeadService : ILeadService
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
            "LeadInfo"
            , RequestFormat = WebMessageFormat.Json, ResponseFormat = WebMessageFormat.Json, BodyStyle = WebMessageBodyStyle.Wrapped)]
        public string LeadInfo(
            Int32 LeadId,
            Byte Type,
            String LeadDate,
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
            myLog.AppendLine("Execute LeadInfo web method.");
            try
            {
                myLog.AppendLine("Parameters retrieved: ");
                myLog.AppendLine("Id: " + Id.ToString());
                myLog.AppendLine("Type: " + Type.ToString());
                myLog.AppendLine("LeadDate: " + LeadDate);
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
                        ErrorMessage = "Can't Authenticate. Please validate the credentials. Message: " + e.Message
                    };

                    string jsonResult = JsonConvert.SerializeObject(refInfo);
                    return jsonResult;
                }

                #region Search for existing Lead in crm system
                myLog.AppendLine("Search for existing Lead.");
                var fetchXml =
                        @" 
                            <fetch mapping='logical' distinct='true'>
                              <entity name='contact'>
                                <attribute name='contactid' />
                                <attribute name='createdon' />
                                <attribute name='createdby' />
                                <attribute name='lastname' />
                                    <filter type='and'>
                                      <condition attribute='contactid' operator='eq' value='" + Id + @"' />
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
                    myLog.AppendLine("Lead already exists");
                    #region Send Reject Dupe to 3rd party service
                    myLog.AppendLine("Send Reject Dupe to 3rd party Lead service.");

                    var httpWebRequest = (HttpWebRequest)WebRequest.Create("https://www.somewebservice.com/api/RejectDupeLead");
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = "POST";
                    httpWebRequest.Headers.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.ASCII.GetBytes(authInfo)));

                    using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                    {
                        string json = "[{\"Id\":\"" + Id + "\"," +
                                      "\"Organization\":\"" + organization + "\"}]";

                        streamWriter.Write(json);
                        streamWriter.Flush();
                        streamWriter.Close();

                        var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                        using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                        {
                            var result = streamReader.ReadToEnd();
                            myLog.AppendLine("RejectDupeLead result: " + result.ToString());
                            returnedRejectDupeResult = result.ToString();
                        }
                    }
                    #endregion
                }
                else
                {
                    #region Send Accept to 3rd party Lead service
                    myLog.AppendLine("Send Accept to 3rd party Lead service.");
                    var httpWebRequest = (HttpWebRequest)WebRequest.Create("https://www.somewebservice.com/api/AcceptLead");
                    httpWebRequest.ContentType = "application/json";
                    httpWebRequest.Method = "POST";
                    httpWebRequest.Headers.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.ASCII.GetBytes(authInfo)));
                    using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                    {
                        string json = "[{\"Id\":\"" + Id + "\"," +
                            "\"Organization\":\"" + organization + "\"}]";

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
                                ErrorMessage += "Error received from 3rd party service. ";
                            }
                        }

                        if (returnedErrorDetails == String.Empty)
                        {
                           //create records in crm system
                        }
                    }
                    #endregion
                }

                #region Log any errors
                if (returnedRejectDupeResult != String.Empty || returnedErrorDetails != String.Empty || ErrorMessage != String.Empty)
                {
                    Entity note = new Entity(Annotation.EntityLogicalName);
                    note["subject"] = "Error(s) for Id: " + Id.ToString();
                    note["notetext"] = "Detail Logs: " + myLog.ToString();
                    if (returnedErrorDetails != String.Empty)
                        note["notetext"] += "\r\n Accept Error: " + returnedErrorDetails;
                    if (returnedRejectDupeResult != String.Empty)
                        note["notetext"] += "\r\n RejectDuplicate Result: " + returnedRejectDupeResult;
                    if (Error != String.Empty)
                        note["notetext"] += "\r\n Error: " + Error;
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
                string json = JsonConvert.SerializeObject(refInfoResults);
                return json;
            }
            catch (Exception e)
            {
                Results refInfoResults = new Results
                {
                    Result = myLog.ToString(),
                    Error = "Message: " + e.Message + "Error: " + ErrorMessage
                };

                //Log service errors
                Entity note = new Entity(Annotation.EntityLogicalName);
                note["subject"] = "Targeted Info Error for Id: " + Id.ToString();
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


        public class Lead
        {
            public int Id { get; set; }
            public byte Type { get; set; }
            public string LeadDate { get; set; }
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

            [JsonProperty("LeadId")]
            public int LeadId { get; set; }

            [JsonProperty("Lead")]
            public Lead Lead { get; set; }
        }

        public class Results
        {
            public string Result { get; set; }
            public string Error { get; set; }
        }
    }
}

