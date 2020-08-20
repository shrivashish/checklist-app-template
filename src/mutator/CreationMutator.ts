import { mutator } from "satcheljs";
import {
  addChoice,
  deleteChoice,
  showBlankTitleError,
  changeItemCheckedStatus,
  updateChoiceText,
  updateTitle,
  setSettings,
  setContext,
  setAppInitialized,
  setSendingFlag,
  goToPage,
} from "../actions/CreationActions";
import {
  Status,
  ChecklistItem,
  checklistItemState,
  ChecklistViewData,
} from "../utils";
import getStore from "../store/CreationStore";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { ActionSDKUtils } from "../utils/ActionSDKUtils";
import { ISettingsComponentProps } from "../../src/components/SettingsComponent";

mutator(setAppInitialized, (msg) => {
  const store = getStore();
  store.isInitialized = msg.state;
});

mutator(setContext, (msg) => {
  const store = getStore();
  store.context = msg.context;
  if (!ActionSDKUtils.isEmptyObject(store.context.lastSessionData)) {
    const lastSessionData = store.context.lastSessionData;
    const actionInstance: actionSDK.Action = lastSessionData.action;
    const actionInstanceRows = lastSessionData.dataRows;
   
    const itemsCopy: ChecklistItem[] = [];
    if (actionInstanceRows && actionInstanceRows.length > 0) {
      actionInstanceRows.forEach((rowItem, index) => {
        let title = rowItem.columnValues['checklistItem'];
        let state:Status
        if(rowItem.columnValues['status'] == Status.ACTIVE){
         state = Status.ACTIVE;
      }
        if (rowItem.columnValues['status'] == Status.COMPLETED) {
          state = Status.COMPLETED;
        }
        let item: ChecklistItem = new ChecklistItem(
          title,
          state,
          "",
          checklistItemState.MODIFIED,
          "",
          (new Date().getTime() + index).toString(),
          "",
          ""
        );
        itemsCopy.push(item);
      });
      getStore().items = itemsCopy;
    }
    getStore().title = actionInstance.displayName;
  }
});

mutator(addChoice, () => {
  const store = getStore();
  const itemsCopy = [...store.items];
  var item = new ChecklistItem();
  itemsCopy.push(item);
  store.items = itemsCopy;
});

mutator(deleteChoice, (msg) => {
  let item: ChecklistItem = msg.item;
  const store = getStore();
  store.items = store.items.filter((x) => x !== item);
});

mutator(showBlankTitleError, (msg) => {
  let blankTitleError: boolean = msg.blankTitleError;
  const store = getStore();
  store.showBlankTitleError = blankTitleError;
});

mutator(changeItemCheckedStatus, (msg) => {
  let item: ChecklistItem = msg.item;
  let state: boolean = msg.state;
  const store = getStore();
  const itemsCopy = [...store.items];
  var index = itemsCopy.indexOf(item);
  if (index > -1) {
    itemsCopy[index] = itemsCopy[index].clone();
    if (state) {
      itemsCopy[index].status = Status.COMPLETED;
    } else {
      itemsCopy[index].status = Status.ACTIVE;
    }
    itemsCopy[index].itemState = checklistItemState.MODIFIED;
  }
  store.items = itemsCopy;
});

mutator(updateChoiceText, (msg) => {
  let item: ChecklistItem = msg.item;
  let text: string = msg.text;
  const store = getStore();
  const itemsCopy = [...store.items];
  var index = itemsCopy.indexOf(item);
  if (index > -1) {
    itemsCopy[index] = itemsCopy[index].clone();
    itemsCopy[index].title = text;
    itemsCopy[index].itemState = checklistItemState.MODIFIED;
  }
  store.items = itemsCopy;
});

mutator(updateTitle, (msg) => {
  let title: string = msg.title;
  const store = getStore();
  store.showBlankTitleError = false;
  store.title = title;
});

mutator(setSettings, (msg) => {
  let settingProps: ISettingsComponentProps = msg.settingProps;
  const store = getStore();
  store.settings = settingProps;
});

mutator(goToPage, (msg) => {
  const store = getStore();
  store.currentPage = msg.page;
});

mutator(setSendingFlag, (msg) => {
  let value: boolean = msg.value;
  const store = getStore();
  store.isSending = value;
});
