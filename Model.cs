// <copyright file="Model.cs" company="CloudFronts Technologies LLP">
// 2016 CloudFronts Technologies LLP, All Rights Reserved.
// </copyright>
// File Name: Model.cs
// Description: Create an incident from email and remove it from queue item.
// Created: October 30, 2017
// Author: Krishna Bhanushali, CloudFronts Technologies LLP
// Revisions:
// ======================================================================
// VERSION      DATE(mm/dd/yyyy)            Modified By         DESCRIPTION 
// ----------------------------------------------------------------------
// 1.0          10/30/2017                  Krishna Bhanushali         CREATED     
// ======================================================================


namespace Model
{
    using System.Collections.Generic;

    static class CRMMarketingList
    {
        #region Entities
        internal const string LOGICAL_NAME = "list";
        #endregion

        #region Attributes
        internal const string ATTR_MAILCHIMP_MARKETING_LIST_ID = "cf_mailchimpmarketinglistid";
        internal const string ATTR_CREATED_FROM_CODE = "createdfromcode";
        internal const string ATTR_LISTID = "listid";


        #endregion
    }
    static class ListMember
    {
        #region Entities
        internal const string LOGICAL_NAME = "listmember";
        #endregion

        #region Attributes
        internal const string ATTR_ENTITYID = "entityid";
       


        #endregion
    }
    static class MailChimpConfiguration
    {
        #region Entities
        internal const string LOGICAL_NAME = "cf_mailchimpconfiguration";
        #endregion

        #region Attributes
        internal const string ATTR_APIKEY = "cf_apikey";
        internal const string ATTR_MAILCHIMP_CONFIGURATION_ID = "cf_mailchimpconfigurationid";
        internal const string ATTR_MAILCHIMP_PASSWORD = "cf_mailchimppassword";
        internal const string ATTR_MAILCHIMP_USERNAME = "cf_mailchimpusername";
        internal const string ATTR_STATECODE = "statecode";
        internal const string ATTR_MAILCHIMPURL = "cf_mailchimpurl";
        
        #endregion
    }

    static class MailChimpSync
    {
        #region Entities
        internal const string LOGICAL_NAME = "cf_mailchimpsync";
        #endregion

        #region Attributes
        internal const string ATTR_BATCHID = "cf_batchid";
        internal const string ATTR_COMPLETED_AT = "cf_completedat";
        internal const string ATTR_ERRORED_OPERATIONS = "cf_erroredoperations";
        internal const string ATTR_FINISHED_OPERATIONS = "cf_finishedoperations";
        internal const string ATTR_MAILCHIMP_SYNC_ID = "cf_mailchimpsyncid";
        internal const string ATTR_MARKETING_LIST = "cf_marketinglist";
        internal const string ATTR_STATUS = "cf_status";
        internal const string ATTR_SUBMITTED_AT = "cf_submittedat";
        internal const string ATTR_TOTAL_OPERATIONS = "cf_totaloperations";
        #endregion
    }
    public class Operation
    {
        public string method { get; set; }
        public string path { get; set; }
        public string body { get; set; }
    }

    public class MailChimpContactCreateBatchRequest
    {
        public List<Operation> operations { get; set; }
    }
    
    public class Link
    {
        public string rel { get; set; }
        public string href { get; set; }
        public string method { get; set; }
        public string targetSchema { get; set; }
        public string schema { get; set; }
    }

    public class MailChimpContactCreateBatchResponse
    {
        public string id { get; set; }
        public string status { get; set; }
        public int total_operations { get; set; }
        public int finished_operations { get; set; }
        public int errored_operations { get; set; }
        public string submitted_at { get; set; }
        public string completed_at { get; set; }
        public string response_body_url { get; set; }
        public List<Link> _links { get; set; }
    }
}
