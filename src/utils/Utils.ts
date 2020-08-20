import {
  addChecklistItems,
  updateSubtitleText,
} from "../actions/UpdationActions";
import getStore from "../store/UpdationStore";
import {
  Status,
  ChecklistColumnType,
  ChecklistItemRow,
  ChecklistItem,
  checklistItemState,
} from ".";
import { showBlankTitleError } from "../actions/CreationActions";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { ActionSDKUtils } from "./ActionSDKUtils";
import { Localizer } from "./Localizer";
import { ActionSdkHelper } from "../helper/ActionSdkHelper";

export const ADD_ITEM_DIV_ID: string = "add-options-cl";

export function isChecklistExpired() {
  if (
    getStore().actionInstance != null &&
    getStore().actionInstance.status == actionSDK.ActionStatus.Expired
  ) {
    return true;
  }
  return false;
}

export function isChecklistClosed() {
  if (
    getStore().actionInstance != null &&
    getStore().actionInstance.status == actionSDK.ActionStatus.Closed
  ) {
    return true;
  }
  return false;
}

export function isChecklistCreatedByMe() {
  if (
    getStore().actionInstance != null &&
    getStore().context != null &&
    getStore().actionInstance.creatorId == getStore().context.userId
  ) {
    return true;
  }
  return false;
}

export function getCompletedSubtext(
  profile: actionSDK.SubscriptionMember,
  time: string
) {
  let subtext = "";
  if (
    !ActionSDKUtils.isEmptyObject(profile) &&
    !ActionSDKUtils.isEmptyString(time)
  ) {
    let completionTime = getDateString(parseInt(time));
    subtext = Localizer.getString(
      "CompletedBy",
      profile.displayName,
      completionTime
    );
  }
  return subtext;
}

export function getStatus(row: ChecklistItemRow) {
  let state: Status;
  if (row[ChecklistColumnType.status] === Status.ACTIVE) {
    state = Status.ACTIVE;
  } else if (row[ChecklistColumnType.status] === Status.COMPLETED) {
    state = Status.COMPLETED;
  } else {
    state = Status.DELETED;
  }
  return state;
}

export function shouldFetchUserProfiles(items: actionSDK.ActionDataRow[]) {
  let userIds: string[] = [];
  for (let actionInstanceRow of items) {
    let row: ChecklistItemRow = JSON.parse(
      JSON.stringify(actionInstanceRow.columnValues)
    );
    if (row[ChecklistColumnType.status] === Status.COMPLETED) {
      userIds.push(row[ChecklistColumnType.completionUser]);
    }
  }
  return userIds;
}

export function validateChecklistCreation(
  actionInstance: actionSDK.Action,
  actionInstanceRows: actionSDK.ActionDataRow[]
): boolean {
  if (
    ActionSDKUtils.isEmptyObject(actionInstance) ||
    actionInstance.dataTables[0].dataColumns == null ||
    ActionSDKUtils.isEmptyString(actionInstance.displayName) ||
    actionInstanceRows.length < 0
  ) {
    if (
      ActionSDKUtils.isEmptyObject(actionInstance) ||
      ActionSDKUtils.isEmptyString(actionInstance.displayName)
    ) {
      showBlankTitleError(true);
    }
    return false;
  }
  return true;
}

export function getDateString(expiry: number): string {
  return new Date(expiry).toLocaleDateString(getStore().context.locale, {
    weekday: "short",
    month: "short",
    day: "numeric",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    hour12: true,
  });
}

export function fetchAllActionInstanceRows(
  pageSize: number = 100,
  rows: actionSDK.ActionDataRow[] = [],
  continuationToken: string = null
): Promise<boolean> {
  return new Promise<boolean>(async (resolve, reject) => {
    let getDataRowsRequest = new actionSDK.GetActionDataRows.Request(
      getStore().context.actionId,
      null,
      continuationToken,
      pageSize
    );
    actionSDK
      .executeApi(getDataRowsRequest)
      .then((response: actionSDK.GetActionDataRows.Response) => {
        rows = response.dataRows;
        console.info("GetActionDataRows - Response"+JSON.stringify(response));
        if (response.continuationToken) {
          fetchAllActionInstanceRows(pageSize, rows, response.continuationToken)
            .then((response) => {
              resolve(response);
            })
            .catch((error) => {
              reject(error);
            });
        } else {
          let userIds: string[] = shouldFetchUserProfiles(rows);
          addChecklistItems(rows);
          if (userIds.length > 0) {
            fetchActionInstanceRowsUserDetailsNow(userIds)
              .then((success) => {
                resolve(true);
              })
              .catch((error) => {
                reject(error);
              });
          } else {
            resolve(true);
          }
        }
      })
      .catch((error) => {
        reject(error);
      });
  });
}

export function fetchActionInstanceRowsUserDetailsNow(
  userIds: string[]
): Promise<boolean> {
  return new Promise<boolean>(async (resolve, reject) => {
    try{
      let response = await ActionSdkHelper.getResponderDetails(getStore().context.subscription,userIds);
      if (!ActionSDKUtils.isEmptyObject(response.members)) {
        let subscriptionMembersMap = createMapping(response.members);
        updateSubtitleText(subscriptionMembersMap);
      }
      resolve(true);
      }
    catch(error) {
      reject(error);
    }
  });
}

function createMapping(members: actionSDK.SubscriptionMember[]) {
  let subscriptionMembersMap = {};
  members.forEach((member) => {
    subscriptionMembersMap[member.id.toString()] = member;
    console.log("Member details : " + member);
  });
  return subscriptionMembersMap;
}

/**
 * Returns true if user has some unsaved changes in the checklist, false otherwise
 */
export function isChecklistDirty() {
  var actionInstanceRows = updateChecklistRows(getStore().context.userId);
  if (
    !ActionSDKUtils.isEmptyObject(actionInstanceRows) &&
    actionInstanceRows.length !== 0
  ) {
    return true;
  }
  return false;
}

/**
 * Get changed checklist rows
 * @param userId
 */
export function updateChecklistRows(userId: string) {
  var actionInstanceRows = [];
  for (var index = 0; index < getStore().items.length; index++) {
    let item: ChecklistItem = getStore().items[index];
    if (
      item.itemState === checklistItemState.MODIFIED &&
      hasItemChanged(item)
    ) {
      if (
        !ActionSDKUtils.isEmptyString(item.title) ||
        !ActionSDKUtils.isEmptyString(item.rowId)
      ) {
        var rowData: ChecklistItemRow = new ChecklistItemRow();
        var actionInstanceRow: actionSDK.ActionDataRow = {
          actionId: getStore().context.actionId,
          columnValues: JSON.parse(JSON.stringify(rowData)),
        };
        if (!ActionSDKUtils.isEmptyString(item.rowId)) {
          actionInstanceRow.id = item.rowId;
          actionInstanceRow.columnValues[
            ChecklistColumnType.creationUser.toString()
          ] = item.creatorUserId;
          // actionInstanceRow.isUpdate = true;
          if (ActionSDKUtils.isEmptyString(item.title)) {
            item.status = Status.DELETED;
            item.title = getOldTitle(item);
          }
        } else {
          // Add creation details if it is not an update
          actionInstanceRow.columnValues[
            ChecklistColumnType.creationUser.toString()
          ] = userId;
        }
        actionInstanceRow.columnValues[
          ChecklistColumnType.creationTime.toString()
        ] = item.creationTime;
        actionInstanceRow.columnValues[
          ChecklistColumnType.checklistItem.toString()
        ] = item.title;
        actionInstanceRow.columnValues[
          ChecklistColumnType.status.toString()
        ] = item.status.toString();
        if (item.status.toString() === Status.COMPLETED) {
          actionInstanceRow.columnValues[
            ChecklistColumnType.completionUser.toString()
          ] = userId;
          let completionTime = ActionSDKUtils.isEmptyString(item.completionTime)
            ? new Date().getTime().toString()
            : item.completionTime;
          actionInstanceRow.columnValues[
            ChecklistColumnType.completionTime.toString()
          ] = completionTime;
        } else if (item.status.toString() === Status.DELETED) {
          actionInstanceRow.columnValues[
            ChecklistColumnType.deletionUser.toString()
          ] = userId;
          actionInstanceRow.columnValues[
            ChecklistColumnType.deletionTime.toString()
          ] = new Date().getTime().toString();
        }
        actionInstanceRow.columnValues[
          ChecklistColumnType.latestEditUser.toString()
        ] = userId;
        actionInstanceRow.columnValues[
          ChecklistColumnType.latestEditTime.toString()
        ] = new Date().getTime().toString();
        actionInstanceRows.push(actionInstanceRow);
      }
    }
  }
  return actionInstanceRows;
}

/**
 * To check if the items received from server either differs in text or status from the current items on the UI
 * @param item
 */
function hasItemChanged(item: ChecklistItem) {
  if (!ActionSDKUtils.isEmptyString(item.rowId)) {
    let actionInstanceRow = getStore().actionInstanceRows.find(
      (x) => x.id === item.rowId
    );
    if (!ActionSDKUtils.isEmptyObject(actionInstanceRow)) {
      let row: ChecklistItemRow = JSON.parse(
        JSON.stringify(actionInstanceRow.columnValues)
      );
      let originalTitle: string = row[ChecklistColumnType.checklistItem];
      let state: Status = getStatus(row);
      if (item.title === originalTitle && item.status === state) {
        return false;
      }
    }
  } else if (item.status === Status.DELETED) {
    // If the row is new and user has added it and deleted it, don't send it to server
    return false;
  }
  return true;
}

/**
 *  Returns the old title fetched via network in the beginning
 * @param item
 */
function getOldTitle(item: ChecklistItem) {
  let actionInstanceRow = getStore().actionInstanceRows.find(
    (x) => x.id === item.rowId
  );
  if (!ActionSDKUtils.isEmptyObject(actionInstanceRow)) {
    let row: ChecklistItemRow = JSON.parse(
      JSON.stringify(actionInstanceRow.columnValues)
    );
    return row[ChecklistColumnType.checklistItem];
  }
}
