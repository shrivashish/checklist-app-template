import {
  closeChecklist,
  checklistCloseAlertOpen,
  deleteChecklist,
  checklistDeleteAlertOpen,
  updateDueDate,
  checklistExpiryChangeAlertOpen,
  setContext,
  fetchActionInstanceRowsUserDetails,
  updateSubtitleText,
  setDownloadingData,
  setSendingFlag,
  saveChangesFailed,
  downloadReportFailed,
  closeChecklistFailed,
  deleteChecklistFailed,
  setIsActionDeleted,
} from "./../actions/UpdationActions";

import { orchestrator } from "satcheljs";
import {
  initialize,
  fetchActionInstance,
  fetchActionInstanceRows,
  addActionInstance,
  updateActionInstance,
  setAppInitialized,
} from "../actions/UpdationActions";
import getStore from "../store/UpdationStore";
import getCreationStore from "../store/CreationStore";
import { fetchAllActionInstanceRows, updateChecklistRows } from "../utils";
import {
  isChecklistClosed,
  isChecklistExpired,
  fetchActionInstanceRowsUserDetailsNow,
} from "../utils/Utils";

import * as actionSDK from "@microsoft/m365-action-sdk";
import { ActionSDKUtils } from "../utils/ActionSDKUtils";
import { ActionError, ActionErrorCode } from "../utils/ActionError";
import { Localizer } from "../utils/Localizer";
import { InitializationState,Constants } from "../../src/components/common";
import { ActionSdkHelper } from "../helper/ActionSdkHelper";

export enum HttpStatusCode {
  Created = 201,
  Unauthorized = 401,
  NotFound = 404,
}

const LOG_TAG = "ChecklistUpdation";

const handleErrorResponse = (error: ActionError) => {
  if (
    error.errorProps &&
    error.errorProps.statusCode == HttpStatusCode.NotFound
  ) {
    setIsActionDeleted(true);
  }
};

orchestrator(initialize,async () => {
  try{
    let context = await ActionSdkHelper.getContext();
    setContext(context);
    Promise.all([
      Localizer.initialize(),
      fetchActionInstanceNow(),
      fetchAllActionInstanceRows(),
    ])
      .then((results) => {
        setAppInitialized(InitializationState.Initialized);
      })
      .catch((error) => {
        setAppInitialized(InitializationState.Failed);
      });
  
}
  catch(error){
    setAppInitialized(InitializationState.Failed);
  }
});

function fetchActionInstanceNow(): Promise<boolean> {
  return new Promise<boolean>(async(resolve, reject) => {
    try {
      let action = await ActionSdkHelper.getActionInstance(getStore().context.actionId);
      addActionInstance(action);
      resolve(true);
      
    } catch (error) {
      handleErrorResponse(error);
      reject(error);
    }
  });
}

orchestrator(fetchActionInstance, fetchActionInstanceNow);

orchestrator(fetchActionInstanceRows, fetchAllActionInstanceRows);

orchestrator(fetchActionInstanceRowsUserDetails, (msg) => {
  fetchActionInstanceRowsUserDetailsNow(msg.userIds);
});

orchestrator(updateActionInstance, async () => {
  let addRows = [];
  let updateRows = [];

  var actionInstanceRows = updateChecklistRows(getStore().context.userId);
  if (
    ActionSDKUtils.isEmptyObject(actionInstanceRows) ||
    actionInstanceRows.length == 0
  ) {
    await ActionSdkHelper.closeCardView();
  } else {
    setSendingFlag(true);
    //Prepare Request arguments
    actionInstanceRows.forEach((row) => {
      if (ActionSDKUtils.isEmptyString(row.id)) {
        row.id = ActionSDKUtils.generateGUID();
        row.createTime = Date.now();
        row.updateTime = Date.now();
        addRows.push(row);
      } else {
        row.updateTime = Date.now();
        updateRows.push(row);
      }
    });

    ActionSDKUtils.announceText(Localizer.getString("SavingChanges"));
        try{
      let addorupdateResponse =  await ActionSdkHelper.addOrUpdateDataRows(addRows,updateRows);
        setSendingFlag(false);
        if (addorupdateResponse.success) {
          ActionSDKUtils.announceText(Localizer.getString("Saved"));
          await ActionSdkHelper.closeCardView();
        } else {
          ActionSDKUtils.announceText(Localizer.getString("Failed"));
          saveChangesFailed(true);
        }
      }
    catch(error) {
        ActionSDKUtils.announceText(Localizer.getString("Failed"));
        setSendingFlag(false);
        saveChangesFailed(true);
        handleErrorResponse(error);
      };
  }
});


orchestrator(setDownloadingData, async (msg) => {
  try{
    if (msg.downloadingData) {
          let downloadDataResponse= await ActionSdkHelper.downloadResponseAsCSV(getStore().context.actionId,
          Localizer.getString("ChecklistResult", getStore().actionInstance.displayName).substring(0, Constants.ACTION_RESULT_FILE_NAME_MAX_LENGTH));
          setDownloadingData(false);
          if (!downloadDataResponse.success) {
            downloadReportFailed(true);
          }
        }
      }
        catch(error) {
          setDownloadingData(false);
          downloadReportFailed(true);
          handleErrorResponse(error);
        }
});


orchestrator(closeChecklist,async(msg) => {
  let addRows = [];
  let updateRows = [];
  // if the checklist has unsaved changes and save first before closing
  var actionInstanceRows = updateChecklistRows(getStore().context.userId);
  if (
    ActionSDKUtils.isEmptyObject(actionInstanceRows) ||
    actionInstanceRows.length == 0
  ) {
    closeChecklistInternal(msg);
  } else {
    //Prepare Request arguments
    actionInstanceRows.forEach((row) => {
      if (ActionSDKUtils.isEmptyString(row.id)) {
        row.id = ActionSDKUtils.generateGUID();
        row.createTime = Date.now();
        row.updateTime = Date.now();
        addRows.push(row);
      } else {
        row.updateTime = Date.now();
        updateRows.push(row);
      }
    });
    ActionSDKUtils.announceText(Localizer.getString("SavingChanges"));
        try{
        let addorupdateResponse=await ActionSdkHelper.addOrUpdateDataRows(addRows,updateRows);
        if (addorupdateResponse.success) {
          closeChecklistInternal(msg);
        } else {
          closeChecklistFailed(true);
        }
    }
      catch(error) {
        closeChecklistFailed(true);
        handleErrorResponse(error);
      }
  }
});

orchestrator(deleteChecklist, async (msg) => {
  try{
  if (msg.deletingChecklist) {
        let deleteResponse=await ActionSdkHelper.deleteActionInstance(getStore().context.actionId);
        checklistDeleteAlertOpen(false);
        await ActionSdkHelper.closeCardView();
        if (!deleteResponse.success) {
          deleteChecklistFailed(true);
        }
      }
    }
      catch(error) {
        deleteChecklistFailed(true);
        handleErrorResponse(error);
      }
  });

orchestrator(updateDueDate,async (actionMessage) => {
  if (actionMessage.updatingDueDate) {
    var actionInstanceUpdateInfo: actionSDK.ActionUpdateInfo = {
      id: getStore().context.actionId,
      version: getStore().actionInstance.version,
      expiryTime: actionMessage.dueDate,
    };
        await ActionSdkHelper.updateActionInstanceStatus(actionInstanceUpdateInfo);
        checklistExpiryChangeAlertOpen(false);
        updateDueDate(getCreationStore().settings.dueDate, false);
        // TODO - intimate user on failure
  }
  });

async function closeChecklistInternal( msg: { closingChecklist: boolean }) {
  try{
  if (msg.closingChecklist) {
    var actionInstanceUpdateInfo: actionSDK.ActionUpdateInfo = {
      id: getStore().context.actionId,
      version: getStore().actionInstance.version,
      status: actionSDK.ActionStatus.Closed,
    };
        let updateActionResponse = await ActionSdkHelper.updateActionInstanceStatus(actionInstanceUpdateInfo);
        checklistCloseAlertOpen(false);
        if (updateActionResponse.success) {
          await ActionSdkHelper.closeCardView();
        } else {
          closeChecklistFailed(true);
        }
      }
    }
      catch(error)  {
        closeChecklistFailed(true);
        handleErrorResponse(error);
      }
  }

