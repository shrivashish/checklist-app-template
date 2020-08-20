import * as React from "react";
import { UxUtils } from "../common";
import { SettingsIcon, Text } from "@fluentui/react-northstar";
import { ResultVisibility } from "../SettingsComponent";
import './SettingsSummaryComponent.scss';
import { ActionSDKUtils } from '../../utils/ActionSDKUtils';
import { Localizer } from '../../utils/Localizer';

export interface ISettingsSummaryProps {
    dueDate?: Date;
    resultVisibility?: ResultVisibility;
    onRef?: (element: HTMLElement) => void;
    onClick?: () => void;
    showDefaultTitle?: boolean;
}

export class SettingsSummaryComponent extends React.Component<ISettingsSummaryProps> {
    isFocused: boolean = false;

    updateSettingsSummary(): string {
        let settingsStrings: string[] = [];
        if (this.props.dueDate) {
            let dueIn: {} = ActionSDKUtils.getTimeRemaining(this.props.dueDate);
            if (dueIn[ActionSDKUtils.YEARS] > 0) {
                settingsStrings.push(Localizer.getString(dueIn[ActionSDKUtils.YEARS] == 1 ? "DueInYear" : "DueInYears", dueIn[ActionSDKUtils.YEARS]));
            }
            else if (dueIn[ActionSDKUtils.MONTHS] > 0) {
                settingsStrings.push(Localizer.getString(dueIn[ActionSDKUtils.MONTHS] == 1 ? "DueInMonth" : "DueInMonths", dueIn[ActionSDKUtils.MONTHS]));
            }
            else if (dueIn[ActionSDKUtils.WEEKS] > 0) {
                settingsStrings.push(Localizer.getString(dueIn[ActionSDKUtils.WEEKS] == 1 ? "DueInWeek" : "DueInWeeks", dueIn[ActionSDKUtils.WEEKS]));
            }
            else if (dueIn[ActionSDKUtils.DAYS] > 0) {
                settingsStrings.push(Localizer.getString(dueIn[ActionSDKUtils.DAYS] == 1 ? "DueInDay" : "DueInDays", dueIn[ActionSDKUtils.DAYS]));
            }
            else if (dueIn[ActionSDKUtils.HOURS] > 0 && dueIn[ActionSDKUtils.MINUTES] > 0) {
                if (dueIn[ActionSDKUtils.HOURS] == 1 && dueIn[ActionSDKUtils.MINUTES] == 1) {
                    settingsStrings.push(Localizer.getString("DueInHourAndMinute", dueIn[ActionSDKUtils.HOURS], dueIn[ActionSDKUtils.MINUTES]));
                } else if (dueIn[ActionSDKUtils.HOURS] == 1) {
                    settingsStrings.push(Localizer.getString("DueInHourAndMinutes", dueIn[ActionSDKUtils.HOURS], dueIn[ActionSDKUtils.MINUTES]));
                } else if (dueIn[ActionSDKUtils.MINUTES] == 1) {
                    settingsStrings.push(Localizer.getString("DueInHoursAndMinute", dueIn[ActionSDKUtils.HOURS], dueIn[ActionSDKUtils.MINUTES]));
                } else {
                    settingsStrings.push(Localizer.getString("DueInHoursAndMinutes", dueIn[ActionSDKUtils.HOURS], dueIn[ActionSDKUtils.MINUTES]));
                }
            }
            else if (dueIn[ActionSDKUtils.HOURS] > 0) {
                settingsStrings.push(Localizer.getString(dueIn[ActionSDKUtils.HOURS] == 1 ? "DueInHour" : "DueInHours", dueIn[ActionSDKUtils.HOURS]));
            }
            else if (dueIn[ActionSDKUtils.MINUTES] > 0) {
                settingsStrings.push(Localizer.getString(dueIn["minutes"] == 1 ? "DueInMinute" : "DueInMinutes", dueIn[ActionSDKUtils.MINUTES]));
            } else {
                settingsStrings.push(Localizer.getString("DueInMinutes", dueIn[ActionSDKUtils.MINUTES]));
            }
        }

        if (this.props.resultVisibility) {
            if (this.props.resultVisibility == ResultVisibility.All) {
                settingsStrings.push(Localizer.getString("ResultsVisibilitySettingsSummaryEveryone"));
            } else {
                settingsStrings.push(Localizer.getString("ResultsVisibilitySettingsSummarySenderOnly"));
            }
        }

        /*if (this.props.notificationSettings) {
            if (this.props.notificationSettings.mode == NotificationSettingMode.None) {
                settingsStrings.push(Localizer.getString("notifyMeNever"));
            } else if (this.props.notificationSettings.mode == NotificationSettingMode.Daily) {
                settingsStrings.push(Localizer.getString("notifyMeOnceADay"));
            } else if (this.props.notificationSettings.mode == NotificationSettingMode.OnRowCreate) {
                settingsStrings.push(Localizer.getString("notifyMeOnEveryUpdate"));
            }
        }*/
        return settingsStrings.join(", ");
    }

    render() {
        return (
            <div className="settings-footer" {...UxUtils.getTabKeyProps()} ref={(element) => {
                if (this.props.onRef) {
                    this.props.onRef(element);
                }
            }} onClick={() => {
                this.props.onClick();
            }}>
                <SettingsIcon className="settings-icon" outline={true} color="brand" />
                <Text content={this.props.showDefaultTitle ? Localizer.getString("Settings") : this.updateSettingsSummary()} size="small" color="brand" />
            </div>);
    }
}
