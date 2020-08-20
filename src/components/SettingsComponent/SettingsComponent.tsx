import * as React from "react";
import "./SettingsComponent.scss";
import { Checkbox, RadioGroup, Flex } from '@fluentui/react-northstar';
import { NotificationSettingMode, ISettingsComponentProps, SettingsSections, ResultVisibility } from "./SettingsCommon";
import { SettingsUtils } from "./SettingsUtils";
import { DateTimePickerView } from "../DateTime";

export class SettingsComponent extends React.PureComponent<ISettingsComponentProps> {
    private settingProps: ISettingsComponentProps;
    constructor(props: ISettingsComponentProps) {
        super(props);
    }

    componentDidMount() {
        if (this.props.onMount) {
            this.props.onMount();
        }
    }

    render() {
        this.settingProps = {
            notificationSettings: this.props.notificationSettings,
            dueDate: this.props.dueDate,
            locale: this.props.locale,
            resultVisibility: this.props.resultVisibility,
            isResponseAnonymous: this.props.isResponseAnonymous,
            isResponseEditable: this.props.isResponseEditable,
            isMultiResponseAllowed: this.props.hasOwnProperty("isMultiResponseAllowed") ? this.props.isMultiResponseAllowed : false,
            strings: this.props.strings
        };
        return (
            <Flex column>
                {SettingsUtils.shouldRenderSection(SettingsSections.DUE_BY, this.props.excludeSections) ? this.renderDueBySection() : null}
                {SettingsUtils.shouldRenderSection(SettingsSections.RESULTS_VISIBILITY, this.props.excludeSections) ? this.renderResultVisibilitySection() : null}
                {SettingsUtils.shouldRenderSection(SettingsSections.NOTIFICATIONS, this.props.excludeSections) ? this.renderNotificationsSection() : null}
                {SettingsUtils.shouldRenderSection(SettingsSections.MULTI_RESPONSE, this.props.excludeSections) ? this.renderResponseOptionsSection() : null}
            </Flex>
        );
    }

    renderDueBySection() {
        if (this.props.renderDueBySection) {
            return this.props.renderDueBySection();
        } else {
            return (
                <Flex className="settings-item-margin" role="group" aria-label={this.getString("dueBy")} column gap="gap.smaller">
                    <label className="settings-item-title">{this.getString("dueBy")}</label>
                    <div className="settings-indentation">
                        <DateTimePickerView showTimePicker
                            minDate={new Date()}
                            locale={this.props.locale}
                            value={new Date(this.props.dueDate)}
                            placeholderDate={this.getString("datePickerPlaceholder")}
                            placeholderTime={this.getString("timePickerPlaceholder")}
                            renderForMobile={this.props.renderForMobile}
                            onSelect={(date: Date) => {
                                this.settingProps.dueDate = date.getTime();
                                this.props.onChange(this.settingProps);
                            }} />
                    </div>
                </Flex>
            );
        }
    }

    renderResultVisibilitySection() {
        if (this.props.renderResultVisibilitySection) {
            return this.props.renderResultVisibilitySection();
        } else {
            return (
                <Flex
                    className="settings-item-margin"
                    role="group"
                    aria-label={this.getString("resultsVisibleTo")}
                    column gap="gap.smaller">
                    <label className="settings-item-title">{this.getString("resultsVisibleTo")}</label>
                    <div className="settings-indentation">
                        <RadioGroup
                            vertical
                            checkedValue={this.settingProps.resultVisibility}
                            items={SettingsUtils.getVisibilityItems(this.getString("resultsVisibleToAll"), this.getString("resultsVisibleToSender"))}
                            onCheckedValueChange={(e, props) => {
                                this.settingProps.resultVisibility = props.value as ResultVisibility;
                                this.props.onChange(this.settingProps);
                            }}
                        />
                    </div>
                </Flex>
            );
        }
    }

    renderNotificationsSection() {
        if (this.props.renderNotificationsSection) {
            return this.props.renderNotificationsSection();
        } else {
            return (
                <Flex
                    className="settings-item-margin"
                    role="group"
                    aria-label={this.getString("notifications")}
                    column gap="gap.smaller">
                    <label className="settings-item-title">{this.getString("notifications")}</label>
                    {
                        <RadioGroup
                            className="settings-indentation"
                            vertical
                            defaultCheckedValue={this.props.notificationSettings.mode}
                            items={SettingsUtils.getNotificationItems(
                                false,
                                this.getString("notificationsAsResponsesAsReceived"),
                                this.getString("notificationsEverydayAt"),
                                this.getString("timePickerPlaceholder"),
                                this.settingProps.notificationSettings.time,
                                (minutes: number) => {
                                    this.settingProps.notificationSettings.time = minutes;
                                    this.props.onChange(this.settingProps);
                                },
                                this.getString("notificationsNever"), this.props.locale)}
                            onCheckedValueChange={(e, props) => {
                                if (props.value == NotificationSettingMode.Daily) {
                                    this.settingProps.notificationSettings.mode = NotificationSettingMode.Daily;
                                } else if (props.value == NotificationSettingMode.None) {
                                    this.settingProps.notificationSettings.mode = NotificationSettingMode.None;
                                } else {
                                    this.settingProps.notificationSettings.mode = NotificationSettingMode.OnRowCreate;
                                }
                                this.props.onChange(this.settingProps);
                            }}
                        />
                    }
                </Flex>
            );
        }
    }

    renderResponseOptionsSection() {
        if (this.props.renderResponseOptionsSection) {
            return this.props.renderResponseOptionsSection();
        } else {
            return (
                <Flex className="settings-item-margin" role="group" aria-label={this.getString("responseOptions")} column gap="gap.small">
                    <label className="settings-item-title">{this.getString("responseOptions")}</label>
                    <Checkbox
                        role="checkbox"
                        className="settings-indentation"
                        checked={this.props.isMultiResponseAllowed}
                        label={this.getString("multipleResponses")}
                        onChange={(e, props) => {
                            this.settingProps.isMultiResponseAllowed = props.checked;
                            this.props.onChange(this.settingProps);
                        }} />
                </Flex>
            );
        }
    }

    getString(key: string): string {
        if (this.props.strings && this.props.strings.hasOwnProperty(key)) {
            return this.props.strings[key];
        }
        return key;
    }
}
