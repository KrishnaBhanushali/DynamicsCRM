// <copyright file="CallMailChimpAPI.cs" company="CloudFronts Technologies LLP">
// 2016 CloudFronts Technologies LLP, All Rights Reserved.
// </copyright>
// File Name: CallMailChimpAPI.cs
// Description: Create an incident from email and remove it from queue item.
// Created: October 30, 2017
// Author: Krishna Bhanushali, CloudFronts Technologies LLP
// Revisions:
// ======================================================================
// VERSION      DATE(mm/dd/yyyy)            Modified By         DESCRIPTION 
// ----------------------------------------------------------------------
// 1.0          10/30/2017                  Krishna Bhanushali         CREATED     
// ======================================================================
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Json;
using System.ServiceModel;
using System.Text;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

[assembly: CLSCompliant(true)]

namespace MailChimpIntegration.Plugins
{
    /// <summary>
    /// Call Mailchimp plugin
    /// </summary>
    public class CallMailChimpAPI : IPlugin
    {
        /// <summary>
        /// Main method
        /// </summary>
        /// <param name="serviceProvider">Service Provider</param>
        public void Execute(IServiceProvider serviceProvider)
        {
            if (serviceProvider == null)
            {
                return;
            }

            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            try
            {
                EntityReference currentRecordReference = (EntityReference)context.InputParameters["Target"];
                if (currentRecordReference.Equals(null) || currentRecordReference.LogicalName != Model.CRMMarketingList.LOGICAL_NAME)
                {
                    tracer.Trace("Does not contain target or logical name is different.");
                    return;
                }

                Entity currentRecord = service.Retrieve(currentRecordReference.LogicalName, currentRecordReference.Id, new ColumnSet(Model.CRMMarketingList.ATTR_MAILCHIMP_MARKETING_LIST_ID, Model.CRMMarketingList.ATTR_CREATED_FROM_CODE));
                string mailChimpMarketingListID = string.Empty;
                if (string.IsNullOrEmpty(currentRecord.GetAttributeValue<string>(Model.CRMMarketingList.ATTR_MAILCHIMP_MARKETING_LIST_ID)))
                {
                    tracer.Trace("Mail chimp marketing list ID is not present");
                    throw new InvalidPluginExecutionException("Mail chimp marketing list ID is not present");
                }
                else
                {
                    mailChimpMarketingListID = currentRecord.GetAttributeValue<string>(Model.CRMMarketingList.ATTR_MAILCHIMP_MARKETING_LIST_ID);
                }

                string targetEntity = this.GetMemberType(currentRecord.GetAttributeValue<OptionSetValue>(Model.CRMMarketingList.ATTR_CREATED_FROM_CODE).Value.ToString());
              
                EntityCollection members = null;
                switch (targetEntity)
                {
                    case "account":
                        tracer.Trace("Marketing List is of type Account and Account dors not contains Email Address. Thus it will not be synched to Mail Chimp");
                        throw new InvalidPluginExecutionException("Marketing List is of type Account and Account dors not contains Email Address. Thus it will not be synched to Mail Chimp");

                    case "contact":
                        tracer.Trace("Retrieve Contact Members");
                        members = this.RetrieveMembers(service, currentRecord, tracer, "contact", "contactid");
                        break;
                    case "lead":
                        tracer.Trace("Retrieve Lead Members");
                        members = this.RetrieveMembers(service, currentRecord, tracer, "lead", "leadid");
                        break;
                    default:
                        break;
                }

                if (members == null)
                {
                    //// No members retrieved
                    tracer.Trace("No members retrieved");

                    throw new InvalidPluginExecutionException("No members retrieved. Kindly add members with Email address present");
                }
                else
                {
                    //// Create a json object for all the members
                    tracer.Trace("Create a json object for all the members");
                    string jsonData = CreateBatchJSON(mailChimpMarketingListID, members, tracer);

                    //// Call to Web Service 
                    tracer.Trace("Call to Web Service ");
                    this.CallBatchCreateWebService(tracer, service, jsonData, currentRecord.Id);

                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException("An error occurred in the Follow up Plugin plug-in.", ex);
            }
            catch (Exception ex)
            {
                tracer.Trace("Message: {0}", ex.ToString());
                throw;
            }
        }

        /// <summary>
        /// Create json string as per the Mail chimp api 
        /// </summary>
        /// <param name="marketingListID">Marketing list id from config record</param>
        /// <param name="members">Entity Collection marketing list members</param>
        /// <returns>json string</returns>
        private static string CreateBatchJSON(string mailChimpMarketingListID, EntityCollection members, ITracingService tracer)
        {
            string jsonData = string.Empty;

            Model.MailChimpContactCreateBatchRequest requestObject = new Model.MailChimpContactCreateBatchRequest();
            List<Model.Operation> operationlist = new List<Model.Operation>();

            foreach (var memberEntity in members.Entities)
            {
                Model.Operation operation = new Model.Operation();
                string subscribed = "subscribed";
                string firstname = string.Empty;
                string lastname = string.Empty;
                if (memberEntity.Contains("firstname"))
                {
                    firstname = memberEntity.Attributes["firstname"].ToString();
                }

                if (memberEntity.Contains("lastname"))
                {
                    lastname = memberEntity.Attributes["lastname"].ToString();
                }
                if (!string.IsNullOrEmpty(firstname))
                {
                    if (!string.IsNullOrEmpty(lastname))
                    {
                        tracer.Trace("First name and last name present");
                        operation.body = "{\"email_address\":\"" + memberEntity.Attributes["emailaddress1"] + "\",\"status\":\"" + subscribed + "\",\"merge_fields\":{\"FNAME\":\"" + firstname + "\",\"LNAME\":\"" + lastname + "\"}}";
                        tracer.Trace("operation.body: " + operation.body);
                    }
                    else
                    {
                        tracer.Trace("First name present");
                        operation.body = "{\"email_address\":\"" + memberEntity.Attributes["emailaddress1"] + "\",\"status\":\"" + subscribed + "\",\"merge_fields\":{\"FNAME\":\"" + firstname + "\"}}";
                        tracer.Trace("operation.body: " + operation.body);
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(lastname))
                    {
                        tracer.Trace(" last name present");
                        operation.body = "{\"email_address\":\"" + memberEntity.Attributes["emailaddress1"] + "\",\"status\":\"" + subscribed + "\",\"merge_fields\":{\"LNAME\":\"" + lastname + "\"}}";
                        tracer.Trace("operation.body: " + operation.body);
                    }
                    else
                    {
                        tracer.Trace("First name and last name not present");
                        operation.body = "{\"email_address\":\"" + memberEntity.Attributes["emailaddress1"] + "\",\"status\":\"" + subscribed + "\"}";
                        tracer.Trace("operation.body: " + operation.body);
                    }
                }

                //// Create request
                tracer.Trace("Create request");
                operation.method = "POST";
                operation.path = "lists/" + mailChimpMarketingListID + "/members";
                tracer.Trace("Operation path: " + operation.path.ToString());
                operationlist.Add(operation);
            }

            requestObject.operations = operationlist;

            jsonData = GetRequestJSON(requestObject);
            return jsonData;
        }

        /// <summary>
        /// Get object from JSON
        /// </summary>
        /// <param name="responseText">Response JSON</param>
        /// <returns>Object response</returns>
        private static Model.MailChimpContactCreateBatchResponse GetInfoFromJSON(string responseText)
        {
            Model.MailChimpContactCreateBatchResponse createBatchResponse = new Model.MailChimpContactCreateBatchResponse();
            MemoryStream ms = new MemoryStream(Encoding.Unicode.GetBytes(responseText));
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(createBatchResponse.GetType());
            createBatchResponse = (Model.MailChimpContactCreateBatchResponse)serializer.ReadObject(ms);
            return createBatchResponse;
        }

        /// <summary>
        /// Convert an object in to JSON string
        /// </summary>
        /// <param name="req">Object Request</param>
        /// <returns>Json String</returns>
        private static string GetRequestJSON(Model.MailChimpContactCreateBatchRequest req)
        {
            string productData = string.Empty;
            DataContractJsonSerializer js = new DataContractJsonSerializer(typeof(Model.MailChimpContactCreateBatchRequest));
            MemoryStream ms = null;
            ms = new MemoryStream();
            js.WriteObject(ms, req);
            ms.Position = 0;
            StreamReader sr = new StreamReader(ms);
            productData = sr.ReadToEnd();
            byte[] data = Encoding.ASCII.GetBytes(productData);
            return productData;
        }

        /// <summary>
        /// Create Mail chimp sync record
        /// </summary>
        /// <param name="createBatchResponse">Response from Web Service call</param>
        /// <param name="tracer">Tracing Service</param>
        /// <param name="service">Iorganization Service</param>
        private static void CreateMailChimpSyncRecord(Model.MailChimpContactCreateBatchResponse createBatchResponse, ITracingService tracer, IOrganizationService service, Guid marketinglistId)
        {
            if (createBatchResponse != null)
            {
                string bactchid = string.Empty;
                string status = string.Empty;
                string submittedAt = string.Empty;
                string completedAt = string.Empty;
                string responseBodyURL = string.Empty;
                Entity mailChimpSync = new Entity(Model.MailChimpSync.LOGICAL_NAME);

                if (!string.IsNullOrEmpty(createBatchResponse.id))
                {
                    bactchid = createBatchResponse.id;
                    mailChimpSync.Attributes[Model.MailChimpSync.ATTR_BATCHID] = createBatchResponse.id;
                }

                if (!string.IsNullOrEmpty(createBatchResponse.status))
                {
                    status = createBatchResponse.status;
                    mailChimpSync.Attributes[Model.MailChimpSync.ATTR_STATUS] = createBatchResponse.status;
                }

                if (!string.IsNullOrEmpty(createBatchResponse.submitted_at))
                {
                    submittedAt = createBatchResponse.submitted_at;
                    //// Convert string in to date format
                    DateTimeFormatInfo usDtfi = new CultureInfo("en-US").DateTimeFormat;
                    DateTime sunmittedDate = (Convert.ToDateTime(submittedAt, usDtfi)).ToUniversalTime();
                    mailChimpSync.Attributes[Model.MailChimpSync.ATTR_SUBMITTED_AT] = sunmittedDate;
                }

                if (!string.IsNullOrEmpty(createBatchResponse.completed_at))
                {
                    completedAt = createBatchResponse.completed_at;
                    //// Convert string in to date format
                    DateTimeFormatInfo usDtfi = new CultureInfo("en-US").DateTimeFormat;
                    DateTime completedDate = Convert.ToDateTime(completedAt, usDtfi).ToUniversalTime();
                    mailChimpSync.Attributes[Model.MailChimpSync.ATTR_COMPLETED_AT] = completedDate;
                }

                if (!string.IsNullOrEmpty(createBatchResponse.response_body_url))
                {
                   /// Do nothing
                }

                mailChimpSync.Attributes[Model.MailChimpSync.ATTR_ERRORED_OPERATIONS] = createBatchResponse.errored_operations;
                mailChimpSync.Attributes[Model.MailChimpSync.ATTR_FINISHED_OPERATIONS] = createBatchResponse.finished_operations;
                mailChimpSync.Attributes[Model.MailChimpSync.ATTR_TOTAL_OPERATIONS] = createBatchResponse.total_operations;
                mailChimpSync.Attributes[Model.MailChimpSync.ATTR_MARKETING_LIST] = new EntityReference(Model.CRMMarketingList.LOGICAL_NAME, marketinglistId);
                service.Create(mailChimpSync);
            }
            else
            {
                throw new InvalidPluginExecutionException("Response is not found");
            }
        }

        /// <summary>
        /// Call Web Service- Batch create request and response captured
        /// </summary>
        /// <param name="tracer">Tracing service</param>
        /// <param name="service">Organizatiob Service</param>
        /// <param name="jsonData">Request Json</param>
        private void CallBatchCreateWebService(ITracingService tracer, IOrganizationService service, string jsonData, Guid marketinglistid)
        {
            string createBatchResponseJSON = string.Empty;
            Entity mailChimpconfig = null;
            mailChimpconfig = this.RetrieveMailChimpConfiguration(tracer, service);
            string username = string.Empty;
            string password = string.Empty;
            string api = string.Empty;
            string basicURL = string.Empty;

            if (mailChimpconfig.Contains(Model.MailChimpConfiguration.ATTR_APIKEY))
            {
                api = mailChimpconfig.Attributes[Model.MailChimpConfiguration.ATTR_APIKEY].ToString();
                tracer.Trace("API key present" + api);
            }
            if (mailChimpconfig.Contains(Model.MailChimpConfiguration.ATTR_MAILCHIMPURL))
            {
                basicURL = mailChimpconfig.Attributes[Model.MailChimpConfiguration.ATTR_MAILCHIMPURL].ToString();
                tracer.Trace("URL present" + basicURL);
            }
            if (mailChimpconfig.Contains(Model.MailChimpConfiguration.ATTR_MAILCHIMP_USERNAME))
            {
                username = mailChimpconfig.Attributes[Model.MailChimpConfiguration.ATTR_MAILCHIMP_USERNAME].ToString();
                tracer.Trace("UserName present" + username);
            }
            if (mailChimpconfig.Contains(Model.MailChimpConfiguration.ATTR_MAILCHIMP_PASSWORD))
            {
                password = mailChimpconfig.Attributes[Model.MailChimpConfiguration.ATTR_MAILCHIMP_PASSWORD].ToString();
                tracer.Trace("Password present" + password);
            }


            using (WebClientEx client = new WebClientEx())
            {
                string authorizationKey = string.Empty;
                authorizationKey = Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes(string.Format(CultureInfo.InvariantCulture, "{0}:{1}", username, api)));
                client.Timeout = 60000;
                client.Headers.Add(HttpRequestHeader.ContentType, "application/json");
                client.Headers.Add(HttpRequestHeader.Authorization, "Basic " + authorizationKey);
                tracer.Trace("jsonData: " + jsonData);
                createBatchResponseJSON = client.UploadString(basicURL, jsonData);
            }

            tracer.Trace("createBatchResponse :" + createBatchResponseJSON);
            Model.MailChimpContactCreateBatchResponse createBatchResponse = GetInfoFromJSON(createBatchResponseJSON);

            CreateMailChimpSyncRecord(createBatchResponse, tracer, service, marketinglistid);
        }

        /// <summary>
        /// Get the type of marketing list
        /// </summary>
        /// <param name="crmNumber">Account- 1, Contact- 2, Lead- 4</param>
        /// <returns>string value of type</returns>
        private string GetMemberType(string crmNumber)
        {
            string type = string.Empty;
            switch (crmNumber)
            {
                case "1":
                    type = "account";
                    break;
                case "2":
                    type = "contact";
                    break;
                case "4":
                    type = "lead";
                    break;
                default:
                    break;
            }

            return type;
        }

        /// <summary>
        /// Retrieve Mail Chimp configuartion record from CRM
        /// </summary>
        /// <param name="tracer">Tracing Service</param>
        /// <param name="service">Organization Service</param>
        /// <returns>Returns entity configuration record</returns>
        private Entity RetrieveMailChimpConfiguration(ITracingService tracer, IOrganizationService service)
        {
            tracer.Trace("Retrieve Configuration Records");
            string fetchXml = @" <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                  <entity name='cf_mailchimpconfiguration'>
                                    <attribute name='cf_mailchimpconfigurationid' />
                                    <attribute name='cf_mailchimpusername' />
                                    <attribute name='createdon' />
                                    <attribute name='statecode' />
                                    <attribute name='cf_mailchimppassword' />
                                    <attribute name='cf_apikey' />
                                    <attribute name='cf_mailchimpurl' />
                                    <order attribute='createdon' descending='true' />
                                    <filter type='and'>
                                      <condition attribute='cf_mailchimpusername' operator='not-null' />
                                      <condition attribute='cf_mailchimppassword' operator='not-null' />
                                      <condition attribute='cf_mailchimpurl' operator='not-null' />
                                      <condition attribute='cf_apikey' operator='not-null' />
                                    </filter>
                                  </entity>
                                </fetch>";
            EntityCollection configurationCollection = null;
            configurationCollection = service.RetrieveMultiple(new FetchExpression(fetchXml));

            if (configurationCollection.Entities.Count == 0)
            {
                tracer.Trace("Mail Chimp Configuration record is not yet created or some field data is missing");
                throw new InvalidPluginExecutionException("Configuration Record is not yet created or some field data is missing. Kindly create correct MailChimp configuration record");
            }
            else
            {
                tracer.Trace("MailChimp Configuration record present in CRM");
                return configurationCollection.Entities[0];
            }
        }



        /// <summary>
        /// Retrive members based on the type of marketing list, contact or lead
        /// </summary>
        /// <param name="service">Organization service</param>
        /// <param name="marketingList">Marketing list entity</param>
        /// <param name="tracer">Tracing Service</param>
        /// <param name="type">Type of marketing list, contact or lead</param>
        /// <returns>Entity collection of retrieved members</returns>
        private EntityCollection RetrieveMembers(IOrganizationService service, Entity marketingList, ITracingService tracer, string type, string typeid)
        {
            string fetchxml = @"<fetch distinct='true' mapping='logical' output-format='xml-platform' version='1.0'>
                                <entity name='" + type + @"'>
                                 <attribute name='firstname'/>
                                    <attribute name='lastname'/>
                                    <attribute name='emailaddress1'/>
                                 <order attribute='createdon' descending='true' />
                                <filter type='and'>
                                      <condition attribute='emailaddress1' operator='not-null' />
                                    </filter>
                                 <link-entity name='listmember' intersect='true' visible='false' to='" + typeid + @"' from='entityid'>
                                  <link-entity name='list' to='listid' from='listid' alias='ag'>
                                   <filter type='and'>
                                     <condition attribute='listid' value='" + marketingList.Id + @"' operator='eq'/>
                                   </filter>
                                  </link-entity>
                                 </link-entity>
                                </entity>
                                </fetch>";
            tracer.Trace("Retrieve List Members");
            EntityCollection listmembersCollection = service.RetrieveMultiple(new FetchExpression(fetchxml));
            //// If Current Opportunity Snapshot is not present, create a new Opportunity Snapshot for stage mentioned
            if (listmembersCollection.Entities.Count == 0)
            {
                return null;
            }
            else
            {
                tracer.Trace("No. of contacts with Email address present in marketing list: " + listmembersCollection.Entities.Count);
                return listmembersCollection;
            }
        }

        /// <summary>
        /// Retrieve Marketing List members
        /// </summary>
        /// <param name="service">Organization Service</param>
        /// <param name="marketingList">Marketing List entity</param>
        /// <returns>Array list </returns>
        private ArrayList RetriveMarketingListMembers(IOrganizationService service, Entity marketingList)
        {
            ArrayList memberGuids = new ArrayList();

            //// Get all the contact members of the marketing list
            QueryByAttribute query = new QueryByAttribute(Model.ListMember.LOGICAL_NAME);

            //// pass the guid of the Static marketing list
            query.AddAttributeValue(Model.CRMMarketingList.ATTR_LISTID, marketingList.Id);
            query.ColumnSet = new ColumnSet(true);
            EntityCollection entityCollection = service.RetrieveMultiple(query);

            foreach (Entity entity in entityCollection.Entities)
            {
                memberGuids.Add(((EntityReference)entity.Attributes[Model.ListMember.ATTR_ENTITYID]).Id);
            }

            //// if list contains more than 5000 records
            while (entityCollection.MoreRecords)
            {
                query.PageInfo.PageNumber += 1;
                query.PageInfo.PagingCookie = entityCollection.PagingCookie;
                entityCollection = service.RetrieveMultiple(query);

                foreach (Entity entity in entityCollection.Entities)
                {
                    memberGuids.Add(((EntityReference)entity.Attributes[Model.ListMember.ATTR_ENTITYID]).Id);
                }
            }
            return memberGuids;
        }

        /// <summary>
        /// Implementing Web client Ex method
        /// </summary>
        private class WebClientEx : WebClient
        {
            /// <summary>
            /// Gets or sets Timeout
            /// </summary>
            public int Timeout { get; set; }

            /// <summary>
            /// Getting Web Request
            /// </summary>
            /// <param name="address">Uri address</param>
            /// <returns>web request</returns>
            protected override WebRequest GetWebRequest(Uri address)
            {
                var request = base.GetWebRequest(address);
                request.Timeout = this.Timeout;
                return request;
            }
        }
    }
}
