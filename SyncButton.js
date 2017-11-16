/// <reference path="MKWebAPI.js" />
// File Name: SyncButton.js
// Description: 
// Created: November 01, 2017
// Author: Krishna Bhanushali, CloudFronts Technologies LLP)
// Revisions:
// ==================================================================================================================
// VERSION         DATE(DD/MM/YY)            Modified By         DESCRIPTION                           Function Modified
// -----------------------------------------------------------------------------------------------------------------
// 1.0               01/11/2017            Krishna Bhanushali      CREATED
// =================================================================================================================

var MailChimpIntegration = {

    callActionOnButtonClick: function () {
        "use strict";
        debugger;
        var entityId = Xrm.Page.data.entity.getId();
        entityId = entityId.replace("{", "").replace("}", "");

        //// GUID Validation
        if (!MailChimpIntegration.isGuid(entityId)) {
            alert("Opportunity ID is Invalid GUID.");
            return;
        }
        alert("Sync Button clicked")
        MailChimpIntegration.invokeBoundAction("lists", entityId, "cf_MailChimpBatchCreateCall", null, MailChimpIntegration.callActionSuccessCallback, MailChimpIntegration.callActionErrorCallback, null);

    },
    callActionSuccessCallback: function (data) {
        "use strict";
        if (data != null || data != undefined) {


            if (MailChimpIntegration.isObject(data)) {
                var result = { value: [] };
                result.value.push(data);
                // callBacks.Success(result);
            }
            else
                //callBacks.Error("No data or Invalid data retrieved");
                alert("No data or Invelid data retrieved")
        }
        else {
            alert("The sync process is started in background.");
        }
    },
    callActionErrorCallback: function (error) {
        "use strict";


        alert(error.message);
    },
    isGuid: function (value) {
        /// <summary>
        /// Checks whether value is type of GUID or not.
        /// </summary>
        /// <param name="value" type="type"></param>
        /// <returns type=""></returns>
        var validGuid = new RegExp("^({|()?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}(}|))?$")
        if (value && typeof value === "string" && validGuid.test(value)) {
            return true;
        }
        return false;
    },
    parseError: function (resp) {
        if (resp && (resp.response || resp.responseText)) {
            var errorObj = JSON.parse(resp.response ? resp.response : resp.responseText);
            if (errorObj.error) {
                return errorObj.error.message;
            }
            if (errorObj.Message) {
                return errorObj.Message;
            }
            return "Unexpected Error";
        }
        return "Unexpected Error";
    },
    errorHandler: function (resp) {
        var errorLog = "";
        switch (resp.status) {
            case 503:
                errorLog = resp.statusText + " Status Code:" + resp.status + " The Web API Preview is not enabled.";
               // console.log(errorLog);
                return new Error(errorLog);

            default:
                errorLog = "Status Code:" + resp.status + " " + MailChimpIntegration.parseError(resp);
               // console.log(errorLog);
                return new Error("Status Code:" + resp.status + " " + MailChimpIntegration.parseError(resp));

        }
    },

    invokeBoundAction: function (entityName, entityId, actionName, parameterObj, successCallback, errorCallback, callerId) {
        /// <summary>Invoke an unbound action</summary>
        /// <param name="actionName" type="String">The name of the unbound action you want to invoke.</param>
        /// <param name="parameterObj" type="Object">An object that defines parameters expected by the action</param>        
        /// <param name="successCallback" type="Function">The function to call when the action is invoked. The results of the action will be passed to this function.</param>
        /// <param name="errorCallback" type="Function">The function to call when there is an error. The error will be passed to this function.</param>
        /// <param name="callerId" type="String" optional="true">The systemuserid value of the user to impersonate</param>
       

        if (actionName.indexOf("Microsoft.Dynamics.CRM") < 0) {
            actionName = "Microsoft.Dynamics.CRM." + actionName
        }
        var uri = MailChimpIntegration.getWebAPIPath() + entityName + "(" + entityId + ")/" + actionName;
        var req = new XMLHttpRequest();
        req.open("POST", encodeURI(uri), true);
        req.setRequestHeader("Accept", "application/json");
        if (callerId) {
            req.setRequestHeader("MSCRMCallerID", callerId);
        }
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.onreadystatechange = function () {
            if (this.readyState == 4 /* complete */) {
                req.onreadystatechange = null;
                if (this.status == 200 || this.status == 201 || this.status == 204) {
                    if (successCallback)
                        switch (this.status) {
                            case 200:
                                //When the Action returns a value
                                successCallback(JSON.parse(this.response, dateReviver));
                                break;
                            case 201:
                            case 204:
                                //When the Action does not return a value
                                successCallback();
                                break;
                            default:
                                //Should not happen
                                break;
                        }

                }
                else {
                    if (errorCallback)
                        errorCallback(MailChimpIntegration.errorHandler(this));
                }
            }
        };
        if (parameterObj) {
            req.send(JSON.stringify(parameterObj));
        }
        else {
            req.send();
        }


    },
    getWebAPIPath: function () {
        return MailChimpIntegration.getClientUrl() + "/api/data/v8.0/";
    },
    getClientUrl: function () {
        return MailChimpIntegration.getContext().getClientUrl();
    },
    isObject: function (value) {
        /// <summary>
        /// Checks whether value is type of Object or not.
        /// </summary>
        /// <param name="value" type="type"></param>
        /// <returns type=""></returns>
        if (value && typeof value === "object") {
            return true;
        }
        return false;
    },
    getContext: function () {
        if (typeof GetGlobalContext != "undefined")
        { return GetGlobalContext(); }
        else {
            if (typeof Xrm != "undefined") {
                return Xrm.Page.context;
            }
            else { throw new Error("Context is not available."); }
        }
    }
};
