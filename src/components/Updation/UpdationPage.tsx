import * as React from "react";
import { observer } from "mobx-react";
import getStore from "../../store/UpdationStore";
import getCreationStore from "../../store/CreationStore";
import {
  updateActionInstance,
  addChoice,
  toggleDeleteChoice,
  updateChoiceText,
  changeItemCheckedStatus,
  checklistCloseAlertOpen,
  closeChecklist,
  checklistDeleteAlertOpen,
  deleteChecklist,
  checklistExpiryChangeAlertOpen,
  updateDueDate,
  setDownloadingData,
  addActionInstance
} from "../../actions/UpdationActions";
import "../../css/updation";
import {
  Flex,
  Dialog,
  Text,
  FlexItem,
  Loader,
  ButtonProps,
  MoreIcon,
  CalendarIcon,
  BanIcon,
  TrashCanIcon,
} from "@fluentui/react-northstar";
import {
  ChecklistGroupType,
  isChecklistExpired
} from "../../utils";
import {ChecklistGroupContainer} from "../ChecklistGroupContainer";
import { setSettings } from "../../actions/CreationActions";
import {
  isChecklistClosed,
  getDateString,
  isChecklistCreatedByMe,
  isChecklistDirty,
} from "../../utils/Utils";

import * as actionSDK from "@microsoft/m365-action-sdk";
import { Localizer } from "../../utils/Localizer";
import { ISettingsComponentProps } from "../SettingsComponent";
import { AdaptiveMenuItem, AdaptiveMenu, AdaptiveMenuRenderStyle } from "../Menu";
import { InitializationState, UxUtils, Constants } from "../common";
import { LoaderUI } from "../Loader";
import { ErrorView } from "../ErrorView";
import { ShimmerContainer } from "../ShimmerLoader";
import { AccessibilityAlert } from "../AccessibilityAlert/AccessibilityAlert";
import { DateTimePickerView } from "../DateTime";
import { ButtonComponent } from "../Button";
import { ActionSdkHelper } from "../../helper/ActionSdkHelper";


@observer
export default class UpdationPage extends React.Component<any, any> {
  render() {
    let hostContext: actionSDK.ActionSdkContext = getStore().context;
    if (hostContext) {
     ActionSdkHelper.hideLoadIndicator();
    } else {
      if (getStore().isInitialized == InitializationState.NotInitialized) {
        return <LoaderUI fill />;
      }
    }

    if (getStore().isActionDeleted) {
     ActionSdkHelper.hideLoadIndicator();
      return (
        <ErrorView
          title={Localizer.getString("ChecklistDeletedError")}
          subtitle={Localizer.getString("ChecklistDeletedErrorDescription")}
          buttonTitle={Localizer.getString("Close")}
          image={"./images/actionDeletedError.png"}
        />
      );
    }

    if (getStore().isInitialized === InitializationState.Failed) {
     ActionSdkHelper.hideLoadIndicator();
      return (
        <ErrorView
          title={Localizer.getString("GenericError")}
          buttonTitle={Localizer.getString("Close")}
        />
      );
    }

    if (getStore().isInitialized == InitializationState.Initialized) {
      ActionSdkHelper.hideLoadIndicator();
    }

    return (
      <>
        <Flex column className="body-container no-mobile-footer">
          {this.getHeaderContainer()}
          {this.getHintText()}
          {this.getItemsGroupSection(ChecklistGroupType.Open)}
          {getStore().isInitialized == InitializationState.Initialized &&
            this.getItemsGroupSection(ChecklistGroupType.Completed)}
        </Flex>
        {
        getStore().isInitialized != InitializationState.Initialized
          ? null
          : this.getFooterSection()}
      </>
    );
  }

  private getHeaderContainer(): JSX.Element {
    return (
      <ShimmerContainer
        fill
        showShimmer={!getStore().actionInstance}
        width={["50%"]}
      >
        <Flex vAlign="center" className={"header-container"}>
          <Text size="large" weight="bold">
            {getStore().actionInstance
              ? getStore().actionInstance.displayName
              : "ChecklistTitle"}
          </Text>
          {this.getMenu()}
          {this.getCloseAlertDialog()}
          {this.getDeleteAlertDialog()}
          {this.getExpiryUpdateDialog()}
        </Flex>
      </ShimmerContainer>
    );
  }

  /* Show due date for open checklist else show "checklist expired/closed" text */
  getHintText() {
    if (isChecklistExpired()) {
      return (
        <Text className="hint-text error" size="small">
          {Localizer.getString("ChecklistExpired")}
        </Text>
      );
    } else if (isChecklistClosed()) {
      return (
        <Text className="hint-text error" size="small">
          {" "}
          {Localizer.getString("ChecklistClosed")}
        </Text>
      );
    }
    let actionInstance: actionSDK.Action = getStore().actionInstance;
    if (actionInstance) {
      let expiry: number = getStore().actionInstance.expiryTime;
      if (expiry != Constants.ACTION_INSTANCE_INDEFINITE_EXPIRY) {
        return (
          <Text className="hint-text" size="small">
            {Localizer.getString("DueByX", getDateString(expiry))}
          </Text>
        );
      }
    }
    return null;
  }

  getItemsGroupSection(checklistGroupType: ChecklistGroupType) {
    return (
      <ChecklistGroupContainer
        sectionType={checklistGroupType}
        items={getStore().items}
        toggleDeleteChoice={(i) => {
          toggleDeleteChoice(i);
        }}
        addChoice={() => {
          addChoice();
        }}
        updateChoiceText={(i, value) => {
          updateChoiceText(i, value);
        }}
        changeItemCheckedStatus={(i, value) => {
          changeItemCheckedStatus(i, value);
        }}
        showShimmer={
          checklistGroupType == ChecklistGroupType.Open &&
          getStore().isInitialized != InitializationState.Initialized
        }
      />
    );
  }

  private getMenu() {
    let menuItems: AdaptiveMenuItem[] = this.getMenuItems();
    if (menuItems.length == 0) {
      return null;
    }
    return (
      <AdaptiveMenu
        key="checklist_options"
        className="triple-dot-menu"
        renderAs={
          UxUtils.renderingForMobile()
            ? AdaptiveMenuRenderStyle.ACTIONSHEET
            : AdaptiveMenuRenderStyle.MENU
        }
        content={
          <MoreIcon
            title={Localizer.getString("MoreOptions")}
            outline
            aria-hidden={false}
            role="button"
          />
        }
        menuItems={menuItems}
        dismissMenuAriaLabel={Localizer.getString("DismissMenu")}
      />
    );
  }

  private getMenuItems(): AdaptiveMenuItem[] {
    let menuItemList: AdaptiveMenuItem[] = [];
    if (isChecklistCreatedByMe()) {
      if (!isChecklistClosed() && !isChecklistExpired()) {
        if (
          getStore().actionInstance.expiryTime !=
          Constants.ACTION_INSTANCE_INDEFINITE_EXPIRY
        ) {
          let changeExpiry: AdaptiveMenuItem = {
            key: "changeDueDate",
            content: Localizer.getString("ChangeDate"),
            icon: <CalendarIcon outline={true} />,
            onClick: () => {
              checklistExpiryChangeAlertOpen(true);
            },
          };
          menuItemList.push(changeExpiry);
        }
        let closeCL: AdaptiveMenuItem = {
          key: "close",
          content: Localizer.getString("CloseChecklist"),
          icon: <BanIcon outline={true} />,
          onClick: () => {
            checklistCloseAlertOpen(true);
          },
        };
        menuItemList.push(closeCL);
      }
      let deleteCL: AdaptiveMenuItem = {
        key: "delete",
        content: Localizer.getString("DeleteChecklist"),
        icon: <TrashCanIcon outline={true} />,
        onClick: () => {
          checklistDeleteAlertOpen(true);
        },
      };
      menuItemList.push(deleteCL);
    }
    return menuItemList;
  }

  getCloseAlertDialog() {
    if (getStore().isChecklistCloseAlertOpen) {
      return (
        <Dialog
          className="dialog-base"
          overlay={{
            className: "dialog-overlay",
          }}
          open={getStore().isChecklistCloseAlertOpen}
          onOpen={(e, { open }) => checklistCloseAlertOpen(open)}
          cancelButton={this.getDialogButtonProps(
            Localizer.getString("CloseChecklist"),
            Localizer.getString("Cancel")
          )}
          confirmButton={
            getStore().closingChecklist && !getStore().closeChecklistFailed ? (
              <Loader size="small" />
            ) : (
              this.getDialogButtonProps(
                Localizer.getString("CloseChecklist"),
                Localizer.getString("Confirm")
              )
            )
          }
          content={
            <Flex gap="gap.smaller" column>
              <Text
                content={
                  isChecklistDirty()
                    ? Localizer.getString("CloseAndSaveAlertDialogMessage")
                    : Localizer.getString("CloseAlertDialogMessage")
                }
              />
              {getStore().closeChecklistFailed ? (
                <Text
                  content={Localizer.getString("SomethingWentWrong")}
                  className="error"
                />
              ) : null}
              {getStore().closeChecklistFailed ? (
                <AccessibilityAlert
                  alertText={Localizer.getString("SomethingWentWrong")}
                />
              ) : null}
            </Flex>
          }
          header={Localizer.getString("CloseChecklist")}
          onCancel={() => {
            checklistCloseAlertOpen(false);
          }}
          onConfirm={() => {
            closeChecklist(true);
          }}
        />
      );
    }
  }

  getDialogButtonProps(
    dialogDescription: string,
    buttonLabel: string
  ): ButtonProps {
    let buttonProps: ButtonProps = {
      content: buttonLabel,
    };

    if (UxUtils.renderingForMobile()) {
      Object.assign(buttonProps, {
        "aria-label": Localizer.getString(
          "DialogTalkback",
          dialogDescription,
          buttonLabel
        ),
      });
    }
    return buttonProps;
  }

  getDeleteAlertDialog() {
    if (getStore().isChecklistDeleteAlertOpen) {
      return (
        <Dialog
          className="dialog-base"
          overlay={{
            className: "dialog-overlay",
          }}
          open={getStore().isChecklistDeleteAlertOpen}
          onOpen={(e, { open }) => checklistDeleteAlertOpen(open)}
          cancelButton={this.getDialogButtonProps(
            Localizer.getString("DeleteChecklist"),
            Localizer.getString("Cancel")
          )}
          confirmButton={
            getStore().deletingChecklist &&
            !getStore().deleteChecklistFailed ? (
              <Loader size="small" />
            ) : (
              this.getDialogButtonProps(
                Localizer.getString("DeleteChecklist"),
                Localizer.getString("Confirm")
              )
            )
          }
          content={
            <Flex gap="gap.smaller" column>
              <Text content={Localizer.getString("DeleteAlertDialogMessage")} />
              {getStore().deleteChecklistFailed ? (
                <Text
                  content={Localizer.getString("SomethingWentWrong")}
                  className="error"
                />
              ) : null}
              {getStore().deleteChecklistFailed ? (
                <AccessibilityAlert
                  alertText={Localizer.getString("SomethingWentWrong")}
                />
              ) : null}
            </Flex>
          }
          header={Localizer.getString("DeleteChecklist")}
          onCancel={() => {
            checklistDeleteAlertOpen(false);
          }}
          onConfirm={() => {
            deleteChecklist(true);
          }}
        />
      );
    }
  }

  getExpiryUpdateDialog() {
    if (getStore().actionInstance && getStore().isChecklistExpiryAlertOpen) {
      return (
        <Dialog
          className="due-date-dialog"
          overlay={{
            className: "dialog-overlay",
          }}
          open={getStore().isChecklistExpiryAlertOpen}
          onOpen={(e, { open }) => checklistExpiryChangeAlertOpen(open)}
          cancelButton={Localizer.getString("Cancel")}
          confirmButton={
            getStore().updatingDueDate ? (
              <Loader size="small" />
            ) : (
              Localizer.getString("Change")
            )
          }
          content={
            <DateTimePickerView
              showTimePicker
              locale={getStore().context.locale}
              renderForMobile={UxUtils.renderingForMobile()}
              minDate={new Date()}
              value={new Date(getStore().actionInstance.expiryTime)}
              placeholderDate={Localizer.getString("SelectADate")}
              placeholderTime={Localizer.getString("SelectATime")}
              onSelect={(date: Date) => {
                let props: ISettingsComponentProps = {
                  ...getCreationStore().settings,
                };
                props.dueDate = date.getTime();
                setSettings(props);
              }}
            />
          }
          header="Action confirmation"
          onCancel={() => {
            checklistExpiryChangeAlertOpen(false);
          }}
          onConfirm={() => {
            updateDueDate(getCreationStore().settings.dueDate, true);
            let actionInstance: actionSDK.Action = {
              ...getStore().actionInstance,
            };
            actionInstance.expiryTime = getCreationStore().settings.dueDate;
            addActionInstance(actionInstance);
          }}
        />
      );
    } else {
      return null;
    }
  }

  getFooterSection() {
    return (
      <Flex className="footer-layout" gap="gap.small">
        {getStore().saveChangesFailed || getStore().downloadReportFailed ? (
          <Text
            content={Localizer.getString("SomethingWentWrong")}
            className="error"
          />
        ) : null}
        {getStore().saveChangesFailed || getStore().downloadReportFailed ? (
          <AccessibilityAlert
            alertText={Localizer.getString("SomethingWentWrong")}
          />
        ) : null}
        <FlexItem push>
          <ButtonComponent
            secondary
            showLoader={getStore().downloadingData}
            content={Localizer.getString("DownloadReport")}
            onClick={() => {
              setDownloadingData(true);
            }}
          />
        </FlexItem>
        <ButtonComponent
          showLoader={getStore().isSending}
          disabled={isChecklistExpired() || isChecklistClosed()}
          primary
          content={Localizer.getString("SaveChanges")}
          onClick={() => {
            updateActionInstance();
          }}
        />
      </Flex>
    );
  }
}
