import * as React from 'react';
import { Stack, Text, List, Icon } from '@fluentui/react';

export interface HeroListItem {
    icon: string;
    primaryText: string;
}

export interface HeroListProps {
    message: string;
    items: HeroListItem[];
    children?: React.ReactNode;
}

export default class HeroList extends React.Component<HeroListProps> {
    private _onRenderCell = (item: HeroListItem | undefined): JSX.Element => {
        if (!item) return <></>;
        return (
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} styles={{ root: { padding: 10 } }}>
                <Icon iconName={item.icon} />
                <Text>{item.primaryText}</Text>
            </Stack>
        );
    };

    render() {
        const {
            children,
            items,
            message,
        } = this.props;

        return (
            <Stack tokens={{ childrenGap: 20 }} styles={{ root: { padding: 20 } }}>
                <Text variant="xLarge">{message}</Text>
                <List items={items} onRenderCell={this._onRenderCell} />
                {children}
            </Stack>
        );
    }
}
