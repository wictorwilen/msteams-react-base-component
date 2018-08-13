# Microsoft Teams UI Controls base component

[![npm version](https://badge.fury.io/js/msteams-react-base-component.svg)](https://badge.fury.io/js/msteams-react-base-component)

This is a base component for Microsoft Teams React based Single Page Applications (SPA), based on the Microsoft Teams UI UI components, which is used when generating Microsoft Teams Apps using the [Microsoft Teams Yeoman Generator](https://aka.ms/yoteams).

 | @master | @preview |
 :--------:|:---------:
 [![Build Status](https://travis-ci.org/wictorwilen/msteams-react-base-component.svg?branch=master)](https://travis-ci.org/wictorwilen/msteams-react-base-component)|[![Build Status](https://travis-ci.org/wictorwilen/msteams-react-base-component.svg?branch=preview)](https://travis-ci.org/wictorwilen/msteams-react-base-component)


## Usage

Example of usage:

```  TypeScript
import * as React from 'react';
import {
    TeamsComponentContext,
    ConnectedComponent,
    Panel,
    PanelBody,
    PanelHeader,
    PanelFooter,
    Surface
} from 'msteams-ui-components-react';
import { render } from 'react-dom';
import TeamsBaseComponent, { ITeamsBaseComponentProps, ITeamsBaseComponentState } from 'msteams-react-base-component'
import * as microsoftTeams from '@microsoft/teams-js';

export interface IMyTabState extends ITeamsBaseComponentState {
    property: string;
}

export interface IMyTabConfigProps extends ITeamsBaseComponentProps {
}

export class MyTab extends TeamsBaseComponent<IMyTapProps, IMyTabState> {
    public componentWillMount() {
        this.updateTheme(this.getQueryVariable('theme'));
        this.setState({
            fontSize: this.pageFontSize()
        });

        microsoftTeams.initialize();
        microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
    }

     public render() {
        return (
            <TeamsComponentContext
                fontSize={this.state.fontSize}
                theme={this.state.theme}>
                <ConnectedComponent render={(props) => {
                    const { context } = props;
                    const { rem, font } = context;
                    const { sizes, weights } = font;
                    const styles = {
                        header: { ...sizes.title, ...weights.semibold },
                        section: { ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4) },
                        footer: { ...sizes.xsmall }
                    }

                    return (
                        <Surface>
                            <Panel>
                                <PanelHeader>
                                    <div style={styles.header}>Hello World</div>
                                </PanelHeader>
                                <PanelBody>
                                    <div style={styles.section}>
                                        HelMy Tab 
                                    </div>
                                </PanelBody>
                                <PanelFooter>
                                    <div style={styles.footer}>
                                        (C) Copyright Myself
                                    </div>
                                </PanelFooter>
                            </Panel>
                        </Surface>
                    );
                }}>
                </ConnectedComponent>
            </TeamsComponentContext>
        );
    }
}

```

## License

Copyright (c) Wictor Wilén. All rights reserved.

Licensed under the MIT license.
