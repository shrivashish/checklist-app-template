import * as React from 'react';
import { Flex, Text, ChevronDownIcon } from '@fluentui/react-northstar';
import { ISettingsComponentProps } from './SettingsCommon';
import { SettingsComponent } from './SettingsComponent';
import { UxUtils } from '../common';
import { Localizer } from '../../utils/Localizer';

export interface ISettingsProps extends ISettingsComponentProps {
    onBack: () => void;
}

export class Settings extends React.PureComponent<ISettingsProps> {

    render() {
        return (
            <Flex className="body-container" column gap="gap.medium">
                <SettingsComponent {...this.props} />
                {this.getBackElement()}
            </Flex>
        );
    }


    private getBackElement() {
        if (!this.props.renderForMobile) {
            return (
                <Flex className="footer-layout" gap={"gap.smaller"}>
                    <Flex vAlign="center" className="pointer-cursor" {...UxUtils.getTabKeyProps()} onClick={() => {
                        this.props.onBack();
                    }} >
                        <ChevronDownIcon rotate={90} xSpacing="after" size="small" />
                        <Text content={Localizer.getString("Back")} />
                    </Flex>
                </Flex>
            );
        }
    }
}