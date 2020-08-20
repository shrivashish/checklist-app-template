// This enum should be in sync with ActionCommon\src\model\Visibility
// We are not using the same enum here so as to not take a dependency on 
// the ActionSDK. We need to ensure that the enum values match exactly.
export enum ResultVisibility {
    All = "All",
    Sender = "Sender"
}

// This enum should be in sync with ActionCommon\src\model\NotificationSettingMode
// We are not using the same enum here so as to not take a dependency on 
// the ActionSDK. We need to ensure that the enum values match exactly.
export enum NotificationSettingMode {
    None = "None",
    Daily = "Daily",
    OnRowCreate = "OnRowCreate",
    OnRowUpdate = "OnRowUpdate"
}

export class NotificationSettings {
    mode: NotificationSettingMode;
    time?: number;
    message?: string;

    constructor(mode: NotificationSettingMode, time: number) {
        this.mode = mode;
        this.time = time;
    }
}

export enum SettingsSections {
    DUE_BY,
    RESULTS_VISIBILITY,
    NOTIFICATIONS,
    MULTI_RESPONSE
}

export interface ISettingsComponentProps {
    notificationSettings: NotificationSettings;
    dueDate: number;
    locale?: string;
    resultVisibility: ResultVisibility;
    isResponseEditable: boolean;
    isResponseAnonymous: boolean;
    renderForMobile?: boolean;
    excludeSections?: SettingsSections[];
    isMultiResponseAllowed?: boolean;
    strings: ISettingsComponentStrings;
    renderDueBySection?: () => React.ReactElement<any>;
    renderResultVisibilitySection?: () => React.ReactElement<any>;
    renderNotificationsSection?: () => React.ReactElement<any>;
    renderResponseOptionsSection?: () => React.ReactElement<any>;
    onChange?: (props: ISettingsComponentProps) => void;
    onMount?: () => void;
}

export interface ISettingsComponentStrings {
    dueBy?: string;
    multipleResponses?: string;
    notifications?: string;
    notificationsNever?: string;
    notificationsAsResponsesAsReceived?: string;
    notificationsEverydayAt?: string;
    responseOptions?: string;
    resultsVisibleTo?: string;
    resultsVisibleToAll?: string;
    resultsVisibleToSender?: string;
    datePickerPlaceholder?: string;
    timePickerPlaceholder?: string;
}