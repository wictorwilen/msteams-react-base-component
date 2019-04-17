# Microsoft Teams UI Controls base component

[![npm version](https://badge.fury.io/js/msteams-react-base-component.svg)](https://www.npmjs.com/package/msteams-react-base-component)
[![npm](https://img.shields.io/npm/dt/msteams-react-base-component.svg)](https://www.npmjs.com/package/msteams-react-base-component)
[![MIT](https://img.shields.io/npm/l/msteams-react-base-component.svg)](https://github.com/wictorwilen/msteams-react-base-component/blob/master/LICENSE.md)
[![GitHub issues](https://img.shields.io/github/issues/wictorwilen/msteams-react-base-component.svg)](https://github.com/wictorwilen/msteams-react-base-component/issues)
[![GitHub closed issues](https://img.shields.io/github/issues-closed/wictorwilen/msteams-react-base-component.svg)](https://github.com/wictorwilen/msteams-react-base-component/issues?q=is%3Aissue+is%3Aclosed) 

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
         const context = getContext({
            baseFontSize: this.state.fontSize,
            style: this.state.theme
        });
        const { rem, font } = context;
        const { sizes, weights } = font;
        const styles = {
            header: { ...sizes.title, ...weights.semibold },
            section: { ...sizes.base, marginTop: rem(1.4), marginBottom: rem(1.4) },
            footer: { ...sizes.xsmall }
        };
        return (
             <TeamsThemeContext.Provider value={context}>
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
            </TeamsThemeContext.Provider>
        );
    }
}

```

## License

Copyright (c) Wictor Wilén. All rights reserved.

Licensed under the MIT license.
