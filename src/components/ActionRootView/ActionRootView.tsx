import * as React from "react";
import "./ActionRootView.scss";
import { Provider, teamsTheme, teamsDarkTheme, teamsHighContrastTheme, ThemePrepared } from '@fluentui/react-northstar'
import * as actionSDK from "@microsoft/m365-action-sdk";
import { ErrorView } from "../ErrorView";
import { ActionSDKUtils } from "../../utils/ActionSDKUtils";
import { ActionSdkHelper } from "../../helper/ActionSdkHelper";

interface IActionRootViewState {
    hostContext: actionSDK.ActionSdkContext;
    stringsInitialized: boolean;
    shouldBlockCreationInMeeting: boolean;
    meetingMemberCount?: number;
}
export class ActionRootView extends React.Component<any, IActionRootViewState> {
    private LOG_TAG = "ActionRootView";

    constructor(props: any) {
        super(props);
        this.state = {
            hostContext: null,
            stringsInitialized: false,
            shouldBlockCreationInMeeting: false,
        };
    }

    async componentWillMount() {
                let context: actionSDK.ActionSdkContext = await ActionSdkHelper.getContext();
                this.setState({
                    hostContext: context,
                });


            }
    

    render() {
        if (!this.state.hostContext) {
            return null;
        }

        document.body.className = this.getClassNames();
        document.body.setAttribute(
            "data-hostclienttype",
            this.state.hostContext.hostClientType
        );

        let isRTL = ActionSDKUtils.isRTL(this.state.hostContext.locale);
        document.body.dir = isRTL ? "rtl" : "ltr";

        ActionSDKUtils.announceText("");

        return (
            <Provider
                theme={this.getTheme()}
                rtl={isRTL}
            >
                {this.props.children}
            </Provider>
        );
    }

    private getTheme(): ThemePrepared {
        switch (this.state.hostContext.theme) {
            case "contrast":
                return teamsHighContrastTheme;

            case "dark":
                return teamsDarkTheme;

            default:
                return teamsTheme;
        }
    }

    private getClassNames(): string {
        let classNames: string[] = [];

        switch (this.state.hostContext.theme) {
            case "contrast":
                classNames.push("theme-contrast");
                break;
            case "dark":
                classNames.push("theme-dark");
                break;
            case "default":
                classNames.push("theme-default");
                break;
            default:
                break;
        }

        if (this.state.hostContext.hostClientType == "android") {
            classNames.push("client-mobile");
            classNames.push("client-android");
        } else if (this.state.hostContext.hostClientType == "ios") {
            classNames.push("client-mobile");
            classNames.push("client-ios");
        } else if (this.state.hostContext.hostClientType == "web") {
            classNames.push("desktop-web");
            classNames.push("web");
        } else if (this.state.hostContext.hostClientType == "desktop") {
            classNames.push("desktop-web");
            classNames.push("desktop");
        } else {
            classNames.push("desktop-web");
        }

        return classNames.join(" ");
    }

    private getUnsupportedPlatformErrorView() {
        // As this is a temporary solution due to Teams Android
        // bug# 3748272 we are not localizing any strings
        let subtitle = "";
        switch (this.state.hostContext.hostClientType) {
            case "android":
                subtitle =
                    "Creation experience is currently not available on Android. Go ahead and use it from your PC";
                break;
            case "ios":
                subtitle =
                    "Creation experience is currently not available on iOS. Go ahead and use it from your PC";
                break;
        }
        return (
            <ErrorView
                image={"./images/unsupportedPlatformError.png"}
                title="Coming Soon!"
                subtitle={subtitle}
                buttonTitle="OK"
            />
        );
    }
}
