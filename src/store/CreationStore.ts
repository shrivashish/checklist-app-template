import { createStore } from "satcheljs";

import { ChecklistItem } from "../utils";
import "../orchestrators/CreationOrchestrators";
import "../mutator/CreationMutator";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { ActionSDKUtils } from "../utils/ActionSDKUtils";
import { ISettingsComponentProps, ResultVisibility, NotificationSettings, NotificationSettingMode } from "../../src/components/SettingsComponent";
import { InitializationState, Constants } from "../../src/components/common";

export enum Page {
  Main,
  Settings,
}

interface IChecklistCreationStore {
  context: actionSDK.ActionSdkContext;
  title: string;
  items: ChecklistItem[];
  settings: ISettingsComponentProps;
  showBlankTitleError: boolean;
  currentPage: Page;
  isInitialized: InitializationState;
  isSending: boolean;
  canChecklistExpire: boolean;
}

const store: IChecklistCreationStore = {
  context: null,
  title: "",
  items: [new ChecklistItem()],
  settings: {
    resultVisibility: ResultVisibility.All,
    dueDate: ActionSDKUtils.getDefaultExpiry(15).getTime(),
    notificationSettings: new NotificationSettings(
      NotificationSettingMode.Daily,
      Constants.DEFAULT_DAILY_NOTIFICATION_TIME
    ),
    isResponseEditable: true,
    isResponseAnonymous: false,
    strings: null,
  },
  showBlankTitleError: false,
  currentPage: Page.Main,
  isInitialized: InitializationState.NotInitialized,
  isSending: false,
  canChecklistExpire: false,
};

export default createStore<IChecklistCreationStore>("creationStore", store);
