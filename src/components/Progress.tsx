import * as React from 'react';
import { Spinner, SpinnerSize, Stack, Text } from '@fluentui/react';

export interface ProgressProps {
    logo: string;
    message: string;
    title: string;
}

export default class Progress extends React.Component<ProgressProps> {
    render() {
        const {
            logo,
            message,
            title,
        } = this.props;

        return (
            <Stack
                horizontalAlign="center"
                verticalAlign="center"
                styles={{
                    root: {
                        height: '100vh',
                        textAlign: 'center'
                    }
                }}
                tokens={{ childrenGap: 15 }}
            >
                <img width='90' height='90' src={logo} alt={title} title={title} />
                <Text variant="xxLarge">{title}</Text>
                <Spinner size={SpinnerSize.large} label={message} />
            </Stack>
        );
    }
}
