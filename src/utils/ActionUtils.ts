import * as actionSDK from "@microsoft/m365-action-sdk";
import { ActionSDKUtils } from "./ActionSDKUtils";

export namespace ActionUtils {
  export function getActionInstanceProperty(
    actionInstance: actionSDK.Action,
    propertyName: string
  ): actionSDK.ActionProperty {
    if (
      actionInstance.customProperties &&
      actionInstance.customProperties.length > 0
    ) {
      for (let property of actionInstance.customProperties) {
        if (property.name == propertyName) {
          return property;
        }
      }
    }
    return null;
  }

  export function prepareActionInstance(
    actionInstance: actionSDK.Action,
    actionContext: actionSDK.ActionSdkContext
  ) {
    if (ActionSDKUtils.isEmptyString(actionInstance.id)) {
      actionInstance.createTime = Date.now();
    }
    actionInstance.updateTime = Date.now();
    actionInstance.creatorId = actionContext.userId;
    actionInstance.actionPackageId = actionContext.actionPackageId;
    actionInstance.version = actionInstance.version || 1;
  }

  export function prepareActionInstanceRow(
    actionInstanceRow: actionSDK.ActionDataRow
  ) {
    if (ActionSDKUtils.isEmptyString(actionInstanceRow.id)) {
      actionInstanceRow.id = ActionSDKUtils.generateGUID();
      actionInstanceRow.createTime = Date.now();
    }
    actionInstanceRow.updateTime = Date.now();
  }

  export function prepareActionInstanceRows(
    actionInstanceRows: actionSDK.ActionDataRow[]
  ) {
    for (let actionInstanceRow of actionInstanceRows) {
      this.prepareActionInstanceRow(actionInstanceRow);
    }
  }
}
