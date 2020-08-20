import * as React from "react";
import { Text, Flex } from '@fluentui/react-northstar';
import { ResultVisibility, NotificationSettingMode, SettingsSections } from "./SettingsCommon";
import { TimePickerView, ITimePickerViewProps } from "../DateTime/TimePickerView";

export class SettingsUtils {
    public static shouldRenderSection(section: SettingsSections, excludedSections: SettingsSections[]) {
        return !excludedSections || (excludedSections.indexOf(section) == -1);
    }

    public static getVisibilityItems(resultsVisibleToAllLabel: string, resultsVisibleToSenderLabel: string) {
        return [
            {
                key: "1",
                label: resultsVisibleToAllLabel,
                value: ResultVisibility.All,
                className: "settings-radio-item"
            },
            {
                key: "2",
                label: resultsVisibleToSenderLabel,
                value: ResultVisibility.Sender,
                className: "settings-radio-item-last"
            },
        ]
    }

    private static adjustLocalTimeinMinutesToUTC(timeinMinutes: number): number {
        let date: Date = new Date();
        date.setHours(timeinMinutes / 60);
        date.setMinutes(timeinMinutes % 60);

        let utcDate: Date = new Date(date.getTime() + date.getTimezoneOffset() * 60000);
        return utcDate.getMinutes() + 60 * utcDate.getHours();
    }

    private static adjustUTCTimeinMinutesToLocal(timeinMinutes: number): number {
        let date: Date = new Date();
        date.setHours(timeinMinutes / 60);
        date.setMinutes(timeinMinutes % 60);

        let localDate: Date = new Date(date.getTime() - date.getTimezoneOffset() * 60000);
        return localDate.getMinutes() + 60 * localDate.getHours();
    }

    public static getNotificationItems(renderForMobile: boolean, notificationsAsResponsesAsReceived: string, notificationsEverydayAt: string, timePickerPlaceholder: string, selectedTime: number, onTimeChange: (time: number) => void, receiveNotification: string, locale: string) {
        let timePickerProps: ITimePickerViewProps = {
            placeholder: timePickerPlaceholder,
            minTimeInMinutes: 0,
            defaultTimeInMinutes: this.adjustUTCTimeinMinutesToLocal(selectedTime) || 0,
            renderForMobile: renderForMobile,
            locale: locale,
            onTimeChange: (minutes: number) => {
                onTimeChange(this.adjustLocalTimeinMinutesToUTC(minutes));
            }
        }
        return [
            {
                key: "1",
                label: (
                    <Flex gap="gap.medium" wrap className="notification-time-picker">
                        <Text content={notificationsEverydayAt} />
                        <TimePickerView {...timePickerProps} />
                    </Flex>
                ),
                className: "settings-radio-item-timepicker",
                value: NotificationSettingMode.Daily
            },
            {
                key: "2",
                label: notificationsAsResponsesAsReceived,
                className: "settings-radio-item",
                value: NotificationSettingMode.OnRowCreate
            },
            {
                key: "3",
                label: receiveNotification,
                className: "settings-radio-item-last",
                value: NotificationSettingMode.None
            }
        ]
    }
}