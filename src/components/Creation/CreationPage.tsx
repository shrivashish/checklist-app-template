import * as React from "react";
import {ChecklistGroupType} from "../../utils";
import {ChecklistItemsContainer} from "../ChecklistItemsContainer";
import {
  callActionInstanceCreationAPI,
  addChoice,
  deleteChoice,
  updateTitle,
  updateChoiceText,
  changeItemCheckedStatus,
  setSettings,
  goToPage
} from "../../actions/CreationActions";
import "../../css/creation";
import getStore, { Page } from "../../store/CreationStore";
import { observer } from "mobx-react";
import {
  Flex,
  Text,
  FlexItem,
  AddIcon,
  ArrowDownIcon,
} from "@fluentui/react-northstar";
import { ADD_ITEM_DIV_ID } from "../../utils/Utils";
import { Localizer } from "../../utils/Localizer";
import { ActionSDKUtils } from "../../utils/ActionSDKUtils";
import * as actionSDK from "@microsoft/m365-action-sdk";
import { ISettingsComponentProps, SettingsSections, ISettingsProps, Settings, ISettingsComponentStrings, SettingsMobile } from "../SettingsComponent";
import { INavBarComponentProps, NavBarComponent, NavBarItemType } from "../NavBarComponent";
import { UxUtils, InitializationState } from "../common";
import { LoaderUI } from "../Loader";
import { ErrorView } from "../ErrorView";
import { InputBox } from "../InputBox";
import { ButtonComponent } from "../Button";
import { ActionSdkHelper } from "../../helper/ActionSdkHelper";


@observer
export default class CreationPage extends React.Component<any, any> {
  private checklistItemsRef;
  private checklistTitleRef: HTMLElement;

  constructor(props) {
    super(props);
    this.checklistItemsRef = React.createRef();
  }

  componentDidUpdate() {
    // If user presses send/create checklist button without filling checklist title, focus should land on title edit field.
    if (getStore().showBlankTitleError && this.checklistTitleRef) {
      this.checklistTitleRef.focus();
    }
  }

  render() {
    if (getStore().isInitialized === InitializationState.NotInitialized) {
      return <LoaderUI fill />;
    } else if (getStore().isInitialized === InitializationState.Failed) {
      ActionSdkHelper.hideLoadIndicator();
      return (
        <ErrorView
          title={Localizer.getString("GenericError")}
          buttonTitle={Localizer.getString("Close")}
        />
      );
    } else {
     ActionSdkHelper.hideLoadIndicator();
      if (UxUtils.renderingForMobile()) {
            
      if (getStore().currentPage == Page.Settings) {
          return this.renderSettingsPageForMobile();
        } else {
          return (
            <Flex className="body-container no-mobile-footer">
              {this.renderChecklistSection()}
              <div className="settings-summary-mobile-container">
                {this.renderFooterSection()}
              </div>
            </Flex>
          );
        }
      } else {
        if (getStore().currentPage == Page.Settings) {
          let settingsProps: ISettingsProps = {
            ...this.getCommonSettingsProps(),
            onBack: () => {
              goToPage(Page.Main);
            },
          };
          return <Settings {...settingsProps} />;
        } else {
          return (
            <>
              <Flex className="body-container" column>
                {this.renderChecklistSection()}
              </Flex>
              {this.renderFooterSection()}
            </>
          );
        }
      }
    }
  }

  renderChecklistSection() {
    let accessibilityAnnouncementString: string = "";
    if (getStore().showBlankTitleError) {
      accessibilityAnnouncementString = Localizer.getString("BlankTitleError");
    }
    ActionSDKUtils.announceText(accessibilityAnnouncementString);
    return (
      <div className="checklist-section">
        <InputBox
          inputRef={(element: HTMLElement) => {
            this.checklistTitleRef = element;
          }}
          fluid
          multiline
          maxLength={240}
          className="title-container"
          input={{
            className: "title-box",
          }}
          defaultValue={getStore().title}
          key="title-box"
          placeholder={Localizer.getString("NameYourChecklist")}
          showError={getStore().showBlankTitleError}
          errorText={
            getStore().showBlankTitleError
              ? Localizer.getString("BlankTitleError")
              : null
          }
          onBlur={(e) => {
            updateTitle((e.target as HTMLInputElement).value);
          }}
        />
        <ChecklistItemsContainer
          ref={(child) => (this.checklistItemsRef = child)}
          sectionType={ChecklistGroupType.All}
          items={getStore().items}
          onToggleDeleteItem={(i) => {
            deleteChoice(i);
          }}
          onItemChecked={(i, value) => {
            changeItemCheckedStatus(i, value);
          }}
          onItemAdded={() => {
            this.onAddChoice();
          }}
          onUpdateItem={(i, value) => {
            updateChoiceText(i, value);
          }}
        />
        <div
          id={ADD_ITEM_DIV_ID}
          className="add-options-cl"
          {...UxUtils.getTabKeyProps()}
          onClick={() => {
            this.onAddChoice();
          }}
        >
          <AddIcon outline size="medium" color="brand" />
          <Text
            className="add-options-cl-label"
            content={Localizer.getString("AddRow")}
            color="brand"
          />
        </div>
        {/* Adding a pseudo-element so that add items button can scroll to the bottom */}
        <div
          id="pseudo-element"
          className="pseudo-element"
          aria-hidden="true"
        />
      </div>
    );
  }

  private onAddChoice() {
    addChoice();
    this.checklistItemsRef.getFocusToLastElement();
    if (!UxUtils.renderingForiOS())
      document.getElementById("pseudo-element").scrollIntoView();
  }

  renderSettingsPageForMobile() {
    let navBarComponentProps: INavBarComponentProps = {
      title: Localizer.getString("Settings"),
      leftNavBarItem: {
        icon: <ArrowDownIcon outline size="large" rotate={90} />,
        ariaLabel: Localizer.getString("Back"),
        onClick: () => {
          goToPage(Page.Main);
        },
        type: NavBarItemType.BACK,
      },
    };

    return (
      <Flex className="body-container no-mobile-footer" column>
        <NavBarComponent {...navBarComponentProps} />
        <SettingsMobile {...this.getCommonSettingsProps()} />
      </Flex>
    );
  }

  //Notification is not enabled
  renderFooterSettingsSection() {
    //if (getStore().context.isNotificationEnabled) {
      return null;
  }

  renderFooterSection() {
    let buttonText: string = Localizer.getString("SendChecklist");

    buttonText = Localizer.getString("Next");
    return (
      <Flex className="footer-layout" gap="gap.small">
        {this.renderFooterSettingsSection()}
        <FlexItem push>
          <ButtonComponent
            primary
            content={buttonText}
            showLoader={getStore().isSending}
            onClick={() => {
              callActionInstanceCreationAPI();
            }}
          />
        </FlexItem>
      </Flex>
    );
  }

  getCommonSettingsProps() {
    let excludeSettingsSections: SettingsSections[] = [
      SettingsSections.MULTI_RESPONSE,
      SettingsSections.RESULTS_VISIBILITY,
      SettingsSections.DUE_BY,
    ];
    /*  if (!getStore().context.isNotificationEnabled) {
      excludeSettingsSections.push(SettingsSections.NOTIFICATIONS);
    }*/
    excludeSettingsSections.push(SettingsSections.NOTIFICATIONS);
    return {
      notificationSettings: getStore().settings.notificationSettings,
      resultVisibility: getStore().settings.resultVisibility,
      isResponseAnonymous: getStore().settings.isResponseAnonymous,
      isResponseEditable: getStore().settings.isResponseAnonymous,
      dueDate: getStore().settings.dueDate,
      locale: getStore().context.locale,
      renderForMobile: UxUtils.renderingForMobile(),
      excludeSections: excludeSettingsSections,
      strings: this.getStringsForSettings(),
      onChange: (props: ISettingsComponentProps) => {
        setSettings(props);
      },
    };
  }

  /**
   * Strings are not there, please add if enabling settings
   */
  getStringsForSettings(): ISettingsComponentStrings {
    let settingsComponentStrings: ISettingsComponentStrings = {
      notifications: Localizer.getString("notifications"),
      notificationsAsResponsesAsReceived: Localizer.getString(
        "notificationsAsResponsesAsReceived"
      ),
      notificationsEverydayAt: Localizer.getString("notificationsEverydayAt"),
      notificationsNever: Localizer.getString("notificationsNever"),
    };
    return settingsComponentStrings;
  }
}
