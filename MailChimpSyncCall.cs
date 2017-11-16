// <copyright file="MailChimpSyncCall.cs" company="CloudFronts Technologies LLP">
// 2016 CloudFronts Technologies LLP, All Rights Reserved.
// </copyright>
// File Name: MailChimpSyncCall.cs
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

namespace MailChimpIntegration.Plugins
{

    public class MailChimpSyncCall : IPlugin
    {
       private IPluginExecutionContext context = null;
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
          context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            try
            {
                EntityReference currentRecordReference = (EntityReference)context.InputParameters["Target"];
                if (currentRecordReference.Equals(null) || currentRecordReference.LogicalName != Model.MailChimpSync.LOGICAL_NAME)
                {
                    tracer.Trace("target record: "+currentRecordReference.LogicalName.ToString());
                    tracer.Trace("Does not contain target or logical name is different.");
                    return;
                }

                //// Retrieve the mail chimp sync record
                Entity currentRecord = service.Retrieve(currentRecordReference.LogicalName, currentRecordReference.Id, new ColumnSet(Model.MailChimpSync.ATTR_BATCHID, Model.MailChimpSync.ATTR_MARKETING_LIST));
                string batchId = string.Empty;
                Guid marketingListId = Guid.Empty;
                if (string.IsNullOrEmpty(currentRecord.GetAttributeValue<string>(Model.MailChimpSync.ATTR_BATCHID)))
                {
                    tracer.Trace("Mail chimp Sync Batch ID is not present");
                    throw new InvalidPluginExecutionException("Mail chimp Sync Batch ID is not present");
                }
                else
                {
                    batchId = currentRecord.GetAttributeValue<string>(Model.MailChimpSync.ATTR_BATCHID);
                }

                if ((currentRecord.GetAttributeValue<EntityReference>(Model.MailChimpSync.ATTR_MARKETING_LIST)) != null)
                {
                    marketingListId = currentRecord.GetAttributeValue<EntityReference>(Model.MailChimpSync.ATTR_MARKETING_LIST).Id;
                }

                string getBatchResponseJSON = string.Empty;
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

                //// Call the web service
                using (WebClientEx client = new WebClientEx())
                {

                  
                    string authorizationKey = string.Empty;
                    authorizationKey = Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes(string.Format(CultureInfo.InvariantCulture, "{0}:{1}", username, api)));
                    
                    basicURL = basicURL + "/" + batchId;

                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(basicURL);
                    request.Accept = "application/json";
                    request.Method = "GET";
                    request.Headers.Add("Authorization", "Basic " + authorizationKey);

                    using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    using (Stream stream = response.GetResponseStream())
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        getBatchResponseJSON = reader.ReadToEnd();
                    }

                }
                tracer.Trace("createBatchResponse :" + getBatchResponseJSON);
                Model.MailChimpContactCreateBatchResponse createBatchResponse = GetInfoFromJSON(getBatchResponseJSON);

                UpdateMailChimpSyncRecord(createBatchResponse, tracer, service, currentRecord.Id);
               

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
        /// Create Mail chimp sync record
        /// </summary>
        /// <param name="updateBatchResponse">Response from Web Service call</param>
        /// <param name="tracer">Tracing Service</param>
        /// <param name="service">Iorganization Service</param>
        private  void UpdateMailChimpSyncRecord(Model.MailChimpContactCreateBatchResponse updateBatchResponse, ITracingService tracer, IOrganizationService service, Guid mailchimpSynctId)
        {
            if (updateBatchResponse != null)
            {
                string bactchid = string.Empty;
                string status = string.Empty;
                string submittedAt = string.Empty;
                string completedAt = string.Empty;
                string responseBodyURL = string.Empty;
                Entity mailChimpSync = new Entity(Model.MailChimpSync.LOGICAL_NAME, mailchimpSynctId);

                if (!string.IsNullOrEmpty(updateBatchResponse.id))
                {
                    bactchid = updateBatchResponse.id;
                    mailChimpSync.Attributes[Model.MailChimpSync.ATTR_BATCHID] = updateBatchResponse.id;
                }

                if (!string.IsNullOrEmpty(updateBatchResponse.status))
                {
                    status = updateBatchResponse.status;
                    mailChimpSync.Attributes[Model.MailChimpSync.ATTR_STATUS] = updateBatchResponse.status;
                }

                if (!string.IsNullOrEmpty(updateBatchResponse.submitted_at))
                {
                    submittedAt = updateBatchResponse.submitted_at;
                    //// Convert string in to date format
                    DateTimeFormatInfo usDtfi = new CultureInfo("en-US").DateTimeFormat;
                    DateTime sunmittedDate = (Convert.ToDateTime(submittedAt, usDtfi)).ToUniversalTime();
                    mailChimpSync.Attributes[Model.MailChimpSync.ATTR_SUBMITTED_AT] = sunmittedDate;
                }

                if (!string.IsNullOrEmpty(updateBatchResponse.completed_at))
                {
                    completedAt = updateBatchResponse.completed_at;
                    //// Convert string in to date format
                    DateTimeFormatInfo usDtfi = new CultureInfo("en-US").DateTimeFormat;
                    DateTime completedDate = Convert.ToDateTime(completedAt, usDtfi).ToUniversalTime();
                    mailChimpSync.Attributes[Model.MailChimpSync.ATTR_COMPLETED_AT] = completedDate;
                }

                if (!string.IsNullOrEmpty(updateBatchResponse.response_body_url))
                {
                   //// Do nothing
                }

                mailChimpSync.Attributes[Model.MailChimpSync.ATTR_ERRORED_OPERATIONS] = updateBatchResponse.errored_operations;
                mailChimpSync.Attributes[Model.MailChimpSync.ATTR_FINISHED_OPERATIONS] = updateBatchResponse.finished_operations;
                mailChimpSync.Attributes[Model.MailChimpSync.ATTR_TOTAL_OPERATIONS] = updateBatchResponse.total_operations;
                service.Update(mailChimpSync);
                tracer.Trace("Mail Chimp sync record is updated with batch Id: " + bactchid);
                context.OutputParameters["MailChimpStatus"] = updateBatchResponse.status;
            }
            else
            {
                throw new InvalidPluginExecutionException("Response is not found");
            }
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
                                  </entity>
                                </fetch>";
            EntityCollection configurationCollection = null;
            configurationCollection = service.RetrieveMultiple(new FetchExpression(fetchXml));

            if (configurationCollection.Entities.Count == 0)
            {
                tracer.Trace("Mail Chimp Configuration record is not yet created");
                throw new InvalidPluginExecutionException("Configuration Record is not yet created. Kindly create MailChimp configuration record");
            }
            else
            {
                tracer.Trace("MailChimp Configuration record present in CRM");
                tracer.Trace("Return True");
                return configurationCollection.Entities[0];
            }
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
