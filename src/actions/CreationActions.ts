import { action } from "satcheljs";
import { ChecklistItem } from "../utils";
import { Page } from "../store/CreationStore";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { ISettingsComponentProps } from "../../src/components/SettingsComponent";
import { InitializationState } from "../../src/components/common";

export enum ChecklistCreationAction {
  initialize = "initialize",
  setContext = "setContext",
  callActionInstanceCreationAPI = "callActionInstanceCreationAPI",
  addChoice = "addChoice",
  deleteChoice = "deleteChoice",
  updateChoiceText = "updateChoiceText",
  changeItemCheckedStatus = "changeItemCheckedStatus",
  updateTitle = "updateTitle",
  goToPage = "goToPage",
  setSettings = "setSettings",
  showBlankTitleError = "showBlankTitleError",
  setAppInitialized = "setAppInitialized",
  setSendingFlag = "setSendingFlag"
}

export let initialize = action(ChecklistCreationAction.initialize);

export let setContext = action(
  ChecklistCreationAction.setContext,
  (context: actionSDK.ActionSdkContext) => ({ context: context })
);

export let callActionInstanceCreationAPI = action(
  ChecklistCreationAction.callActionInstanceCreationAPI
);

export let addChoice = action(ChecklistCreationAction.addChoice);

export let deleteChoice = action(
  ChecklistCreationAction.deleteChoice,
  (item: ChecklistItem) => ({ item: item })
);

export let showBlankTitleError = action(
  ChecklistCreationAction.showBlankTitleError,
  (blankTitleError: boolean) => ({ blankTitleError: blankTitleError })
);

export let updateTitle = action(
  ChecklistCreationAction.updateTitle,
  (title: string) => ({ title: title })
);

export let updateChoiceText = action(
  ChecklistCreationAction.updateChoiceText,
  (item: ChecklistItem, text: string) => ({ item: item, text: text })
);

export let changeItemCheckedStatus = action(
  ChecklistCreationAction.changeItemCheckedStatus,
  (item: ChecklistItem, state: boolean) => ({ item: item, state: state })
);

export let goToPage = action(
  ChecklistCreationAction.goToPage,
  (page: Page) => ({ page: page })
);

export let setSettings = action(
  ChecklistCreationAction.setSettings,
  (props: ISettingsComponentProps) => ({ settingProps: props })
);

export let setAppInitialized = action(
  ChecklistCreationAction.setAppInitialized,
  (state: InitializationState) => ({ state: state })
);

export let setSendingFlag = action(
  ChecklistCreationAction.setSendingFlag,
  (value: boolean) => ({ value: value })
);


