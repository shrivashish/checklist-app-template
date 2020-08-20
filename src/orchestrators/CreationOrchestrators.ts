import { orchestrator } from "satcheljs";
import getStore, { Page } from "../store/CreationStore";
import {
  ChecklistColumnType,
  Status,
  ChecklistItemRow,
  checklistItemState,
  validateChecklistCreation,
  ChecklistItem,
  ChecklistViewData,
} from "../utils";
import {
  initialize,
  callActionInstanceCreationAPI,
  setContext,
  setAppInitialized,
  setSendingFlag,
  goToPage
} from "../actions/CreationActions";

import { toJS } from "mobx";

import * as actionSDK from "@microsoft/m365-action-sdk";
import { Localizer } from "../utils/Localizer";
import { ActionSDKUtils } from "../utils/ActionSDKUtils";
import { ActionUtils } from "../utils/ActionUtils";
import { InitializationState, Constants } from "../../src/components/common";
import { ActionSdkHelper } from "../helper/ActionSdkHelper"

let batchReq = [];

orchestrator(initialize, () => {
  Localizer.initialize()
    .then(async () => {
      try{
          let context=await ActionSdkHelper.getContext();
          setContext(context);
          setAppInitialized(InitializationState.Initialized);
      }
        catch(error) {
          setAppInitialized(InitializationState.Failed);
        };
    })
    .catch((error) => {
      setAppInitialized(InitializationState.Failed);
    });
});

orchestrator(callActionInstanceCreationAPI, () => {
  goToPage(Page.Main);
  var actionInstance: actionSDK.Action = {
    id: ActionSDKUtils.generateGUID(),
    displayName: getStore().title,
    expiryTime: getStore().canChecklistExpire
      ? getStore().settings.dueDate
      : Constants.ACTION_INSTANCE_INDEFINITE_EXPIRY,
    customProperties: [],
    permissions: {
      [actionSDK.ActionPermission.DataRowsUpdate]: [
        actionSDK.ActionRole.Member,
      ],
    },
    dataTables: [
      {
        name: "TestDataSet",
        rowsVisibility: actionSDK.Visibility.All,
        rowsEditable: true,
        canUserAddMultipleRows: true,
        dataColumns: [],
        attachments: [],
      },
    ],
  };

  createChecklistColumns(actionInstance);
  let actionInstanceRows = createChecklistRows(
    getStore().context.userId,
    actionInstance.id
  );
  let viewData = createChecklistViewData();
  if (validateChecklistCreation(actionInstance, actionInstanceRows)) {
    setSendingFlag(true);
    ActionUtils.prepareActionInstance(actionInstance, toJS(getStore().context));
    ActionUtils.prepareActionInstanceRows(actionInstanceRows);
    //Create Action
    let createAction = new actionSDK.CreateAction.Request(actionInstance);
    batchReq.push(createAction);
    console.info("CreateAction - Request: " + JSON.stringify(actionInstance));
    //AddorUpdateRows
    if (
      !ActionSDKUtils.isEmptyObject(actionInstanceRows) &&
      actionInstanceRows.length > 0
    ) {
      let addOrUpdateRowsRequest = new actionSDK.AddOrUpdateActionDataRows.Request(
        actionInstanceRows,
        []
      );
      batchReq.push(addOrUpdateRowsRequest);
    }
    ActionSdkHelper.executeBatchRequest(batchReq);
  }
});

function createChecklistRows(userId: string, actionInstanceId) {
  var actionInstanceRows = [];
  for (var index = 0; index < getStore().items.length; index++) {
    // Only add modified items
    var item = getStore().items[index];
    if (
      item.itemState == checklistItemState.MODIFIED &&
      !ActionSDKUtils.isEmptyString(item.title)
    ) {
      var rowData: ChecklistItemRow = new ChecklistItemRow();
      var actionInstanceRow: actionSDK.ActionDataRow = {
        actionId: actionInstanceId,
        columnValues: JSON.parse(JSON.stringify(rowData)),
      };
      actionInstanceRow.columnValues[
        ChecklistColumnType.checklistItem.toString()
      ] = item.title;
      actionInstanceRow.columnValues[
        ChecklistColumnType.status.toString()
      ] = item.status.toString();
      actionInstanceRow.columnValues[
        ChecklistColumnType.creationUser.toString()
      ] = userId;
      actionInstanceRow.columnValues[
        ChecklistColumnType.creationTime.toString()
      ] = item.creationTime;
      if (item.status.toString() === Status.COMPLETED) {
        actionInstanceRow.columnValues[
          ChecklistColumnType.completionUser.toString()
        ] = userId;
        actionInstanceRow.columnValues[
          ChecklistColumnType.completionTime.toString()
        ] = new Date().getTime().toString();
      }
      actionInstanceRows.push(actionInstanceRow);
    }
  }
  return actionInstanceRows;
}

function createChecklistViewData() {
  let viewData: ChecklistViewData = {
    ait: getStore().title,
    air: [],
  };
  if (
    !ActionSDKUtils.isEmptyObject(getStore().settings.notificationSettings.mode)
  ) {
    viewData.nst = getStore().settings.notificationSettings.time;
  }
  //sort items in order of creation time before sending
  getStore().items.sort((a: ChecklistItem, b: ChecklistItem) => {
    return a.creationTime > b.creationTime
      ? 1
      : b.creationTime > a.creationTime
      ? -1
      : 0;
  });
  for (var index = 0; index < getStore().items.length; index++) {
    var item = getStore().items[index];
    if (
      item.itemState == checklistItemState.MODIFIED &&
      !ActionSDKUtils.isEmptyString(item.title)
    ) {
      let actionInstanceRow: string[] = [];
      actionInstanceRow[0] = item.title;
      //The presence of 2nd param means it is a completed item
      if (item.status === Status.COMPLETED) {
        actionInstanceRow[1] = "1";
      }
      viewData.air.push(actionInstanceRow);
    }
  }
  Object.keys(viewData.air).forEach(
    (key) => viewData.air[key] == null && delete viewData.air[key]
  );
  return viewData;
}

function createChecklistColumns(actionInstance: actionSDK.Action) {
  for (let item in ChecklistColumnType) {
    var checklistCol: actionSDK.ActionDataColumn = {
      name: item,
      valueType: actionSDK.ActionDataColumnValueType.Text,
      displayName: item,
      allowNullValue: true,
    };
    if (
      item.match(ChecklistColumnType.checklistItem) ||
      item.match(ChecklistColumnType.status) ||
      item.match(ChecklistColumnType.creationTime) ||
      item.match(ChecklistColumnType.creationUser)
    ) {
      checklistCol.allowNullValue = false;
    }
    if (item.match(ChecklistColumnType.status)) {
      checklistCol.valueType = actionSDK.ActionDataColumnValueType.SingleOption;
      checklistCol.options = [];
      checklistCol.options.push(statusParams(Status.ACTIVE));
      checklistCol.options.push(statusParams(Status.COMPLETED));
      checklistCol.options.push(statusParams(Status.DELETED));
    }
    if (
      item.match(ChecklistColumnType.completionUser) ||
      item.match(ChecklistColumnType.latestEditUser) ||
      item.match(ChecklistColumnType.creationUser) ||
      item.match(ChecklistColumnType.deletionUser)
    ) {
      // checklistCol.isExcludedFromReporting = true;
      checklistCol.valueType = actionSDK.ActionDataColumnValueType.UserId;
    }
    if (
      item.match(ChecklistColumnType.completionTime) ||
      item.match(ChecklistColumnType.latestEditTime) ||
      item.match(ChecklistColumnType.creationTime) ||
      item.match(ChecklistColumnType.deletionTime)
    ) {
      checklistCol.valueType = actionSDK.ActionDataColumnValueType.DateTime;
    }
    actionInstance.dataTables[0].dataColumns.push(checklistCol);
  }
}

function statusParams(state: Status) {
  var optionActive: actionSDK.ActionDataColumnOption = {
    name: state.toString(),
    displayName: state.toString(),
  };
  return optionActive;
}
