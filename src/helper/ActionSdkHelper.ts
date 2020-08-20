import * as actionSDK from "@microsoft/m365-action-sdk";

export class ActionSdkHelper {

    /*
    * @desc Gets the localized strings in which the app is rendered
    */
    public static async getLocalizedStrings() {
        let request = new actionSDK.GetLocalizedStrings.Request();
        let response = await actionSDK.executeApi(request) as actionSDK.GetLocalizedStrings.Response;
        return response.strings;
    }
    /*
    * @desc Service Request to create new Action Instance 
    * @param {actionSDK.Action} action instance which need to get created
    */
    public static async createActionInstance(action: actionSDK.Action) {
        try {
        let createRequest = new actionSDK.CreateAction.Request(action);
        let createResponse = await  actionSDK.executeApi(createRequest) as  actionSDK.GetContext.Response;
        }
        catch(error) {
            console.error("Error: " + JSON.stringify(error)); //Add error log
        }
    }
    /*
    * @desc Service API Request for fetching action instance
    * @param {actionSDK.ActionSdkContext} - action Sdk context of the App
    */
    public static async getAction(actionId) {
        /*
        * Get Action Instance Details 
        */
        let request = new actionSDK.GetAction.Request(actionId);
        let response = await actionSDK.executeApi(request) as actionSDK.GetAction.Response;
        return response.action;
    }

      /**
     * function to execute batch request
     * @param batchRequestArray Array of request to be executed in batch
     */
    public static async executeBatchRequest(batchRequestArray) {
        let batchRequest = new actionSDK.BaseApi.BatchRequest(batchRequestArray);
        try {
            let batchResponse = await actionSDK.executeBatchApi(batchRequest);
            console.info("BatchResponse: " + JSON.stringify(batchResponse));
            return batchResponse;
        } catch (error) {
            console.log("Console log: Error: " + JSON.stringify(error));
            return;
        }
    }
    
    /*
    * @desc Service API Request for Submit of Response
    * @param {actionSDK.ActionDataRow} - data row of survey response
    */
    public static addDataRows(dataRow: actionSDK.ActionDataRow) {
        let addDataRowRequest = new actionSDK.AddActionDataRow.Request(dataRow);
        let closeViewRequest = new actionSDK.CloseView.Request();
        /*
        * @desc Prepare Batch Request object for simultaneously making multiple APIs Request
        */
        let batchRequest = new actionSDK.BaseApi.BatchRequest([addDataRowRequest, closeViewRequest]);
        actionSDK.executeBatchApi(batchRequest)
            .then(function (batchResponse) {
                console.info("BatchResponse: " + JSON.stringify(batchResponse));
            })
            .catch(function (error) {
                console.error("Error: " + JSON.stringify(error));
            })
    }
    /*
    *   @desc Service API Request for getting the actionContext
    *   @return response.context: actionSDK.ActionSdkContext
    */
    public static async getContext() {
        try {
            let response = await actionSDK.executeApi(new actionSDK.GetContext.Request()) as actionSDK.GetContext.Response;
            return response.context;
        }
        catch(error) {
            console.error("Error: " + JSON.stringify(error)); //Add error log
        }
    }
    /*
    *   @desc Service API Request for getting the actionInstance
    *   @param context - actionInstance context: actionSDK.ActionSdkContext
    *   @return response.action: actionSDK.Action
    */
    public static async getActionInstance(actionId: string) {
        try {
            let getActionRequest = new actionSDK.GetAction.Request(actionId);
            let response = await actionSDK.executeApi(getActionRequest) as actionSDK.GetAction.Response;
            return response.action;
        }
        catch(error) {
            console.error("Error: " + JSON.stringify(error)); //Add error log
        }
    }
    /*
    *   @desc Service API Request for getting the actionInstance response summary
    *   @param context - actionInstance context: actionSDK.ActionSdkContext
    *   @return response.summary: actionSDK.GetActionDataRowsSummary
    */
    public static async getActionSummary(actionId: string) {
        try {
            let getSummaryRequest = new actionSDK.GetActionDataRowsSummary.Request(actionId, true);
            let response = await actionSDK.executeApi(getSummaryRequest) as actionSDK.GetActionDataRowsSummary.Response;
            return response.summary;
        }
        catch(error) {
            console.error("Error: " + JSON.stringify(error)); //Add error log
        }
    }
    /*
    *   @desc Service API Request for getting the actionInstance responses
    *   @param context - actionInstance context: actionSDK.ActionSdkContext
    *   @return response.dataRows: actionSDK.GetActionDataRows
    */
    public static async getActionDataRows(context: actionSDK.ActionSdkContext, creatorId=null, continuationToken=null, pageSize=30, dataTableName=null) {
        try {
            let getDataRowsRequest = new actionSDK.GetActionDataRows.Request(context.actionId, creatorId ,continuationToken, pageSize, dataTableName);
            let response = await actionSDK.executeApi(getDataRowsRequest) as actionSDK.GetActionDataRows.Response;
            return response;
        }
        catch(error) {
            console.error("Error: " + JSON.stringify(error)); //Add error log
        }
    }
    /*
    *   @desc Service API Request for getting the membercount
    *   @param context - action context: actionSDK.ActionSdkContext
    *   @return response.memberCount: number
    */
    public static async getMemberCount(subscription: actionSDK.Subscription) {
        try {
            let getSubscriptionCount = new actionSDK.GetSubscriptionMemberCount.Request(subscription);
            let response = await actionSDK.executeApi(getSubscriptionCount) as actionSDK.GetSubscriptionMemberCount.Response;
            return response.memberCount;
        }
        catch(error) {
            console.error("Error: " + JSON.stringify(error));  //Add error log 
        }
    }
    /*
    *   @desc Service API Request for getting the responders details
    *   @param subscription - actionSDK.Subscription
    *   @param userIds - string array of all the datarows creatorId
    *   @return datarow crestor's details
    */
    public static async getResponderDetails(subscription: actionSDK.Subscription, userIds: string[]) {
        try{
            let requestResponders = new actionSDK.GetSubscriptionMembers.Request(subscription, userIds);
            let responseResponders = await actionSDK.executeApi(requestResponders) as actionSDK.GetSubscriptionMembers.Response;
            return responseResponders;
        }
        catch(error) {
            console.error("Error: " + JSON.stringify(error)); //Add error log
        }
    }

    /*
    *   @desc Service API Request for getting the nonResponders details
    *   @param actionId - context.actionId
    *   @param subscriptionId - context.subscription.id
    *   @return NonResponders: [label:<>,userId:<>]
    */
    public static async getNonResponders(actionId: string, subscriptionId: string) {
        try {
            let requestNonResponders = new actionSDK.GetActionSubscriptionNonParticipants.Request(actionId, subscriptionId);
            let responseNonResponders = await actionSDK.executeApi(requestNonResponders) as actionSDK.GetActionSubscriptionNonParticipants.Response;
            return responseNonResponders;
        }
        catch(error) {
            console.error("Error: " + JSON.stringify(error)); //Add error log
        }
    }
    /*
     * Method to update action instance data 
     * @param data object of data we want modify
    */
    public static async updateActionInstance(actionInstance, data) {
        let action: actionSDK.ActionUpdateInfo = {
            id: actionInstance.id,
            version: actionInstance.version,
            displayName: actionInstance.displayName,
            dataTables: actionInstance.dataTables
        }
        for (let key in data) {
            action[key] = data[key];
        }
        let getUpdateActionRequest = new actionSDK.UpdateAction.Request(action);
        try {
            let response = await actionSDK.executeApi(getUpdateActionRequest) as actionSDK.UpdateAction.Response;
            console.info("UpdateAction - Response: " + JSON.stringify(response));
            actionInstance = await ActionSdkHelper.getActionInstance(actionInstance.id);
            return actionInstance;
        } catch (error) {
            console.error("UpdateAction - Error: " + JSON.stringify(error));
        }
    }
    public static async deleteActionInstance(actionId) {
        try{
            let request = new actionSDK.DeleteAction.Request(actionId);
            let response = await actionSDK.executeApi(request) as actionSDK.DeleteAction.Response;
            return response;
        }
        catch(error) {
            console.error("Error: " + JSON.stringify(error)); //Add error log
        }
            
    }

    public static async downloadResponseAsCSV(actionId: string, fileName: string){
        try {
            let downloadCSVRequest = new actionSDK.DownloadActionDataRowsResult.Request(
                actionId,
                fileName
            );
            let downloadResponse = await actionSDK.executeApi(downloadCSVRequest) as actionSDK.DownloadActionDataRowsResult.Response
            return downloadResponse;
            }
        catch(error) {
            console.error("Error: " + JSON.stringify(error)); //Add error log
        }
    }

    public static async updateActionInstanceStatus(updateInfo){
        try{
            let updateActionRequest = new actionSDK.UpdateAction.Request(updateInfo);
            let updateActionResponse = await actionSDK.executeApi(updateActionRequest) as actionSDK.UpdateAction.Response;
            return updateActionResponse;
        }
        catch(error) {
            console.error("Error: " + JSON.stringify(error)); //Add error log
        }
    }

    public static async closeCardView(){
        try {
            let closeViewRequest = new actionSDK.CloseView.Request();
            await actionSDK.executeApi(closeViewRequest) as actionSDK.CloseView.Response;
        }
        catch(error) {
            console.error("Error: " + JSON.stringify(error)); //Add error log
        }
    }

    public static async addOrUpdateDataRows(addRows, updateRows){
        try {
            let addOrUpdateRowsRequest = new actionSDK.AddOrUpdateActionDataRows.Request(
                addRows,
                updateRows
            );
            let addOrUpdateResponse = await actionSDK.executeApi(addOrUpdateRowsRequest) as actionSDK.AddOrUpdateActionDataRows.Response;
            return addOrUpdateResponse;
        }
        catch(error) {
            console.error("Error: " + JSON.stringify(error)); //Add error log
        }
    }

    public static async hideLoadIndicator(){
        await actionSDK.executeApi(new actionSDK.HideLoadingIndicator.Request());
    }
}