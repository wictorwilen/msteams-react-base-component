// Copyright (c) Wictor Wilén. All rights reserved.
// Licensed under the MIT license.

import * as React from "react";
import { render } from "react-dom";
import { ThemeStyle, ITeamsComponentProps, ITeamsComponentState } from "msteams-ui-components-react";
import * as microsoftTeams from "@microsoft/teams-js";
/**
 * State interface for the Teams Base user interface React component
 */
export interface ITeamsBaseComponentState extends ITeamsComponentState {
    fontSize: number;
    theme: ThemeStyle;
}

/**
 * Properties interface for the Teams Base user interface React component
 */
export interface ITeamsBaseComponentProps extends ITeamsComponentProps {

}

/**
 * Base implementation of the React based interface for the Microsoft Teams app
 */
export default class TeamsBaseComponent<P extends ITeamsBaseComponentProps, S extends ITeamsBaseComponentState>
    extends React.Component<P, S> {

    /**
     * Static method to render the component
     * @param element DOM element to render the control in
     * @param props Properties
     */
    public static render<P extends ITeamsBaseComponentProps>(element: HTMLElement, props: P) {
        render(React.createElement(this, props), element);
    }

    /**
     * Constructor
     * @param props Properties
     * @param state State
     */
    constructor(props: P, state: S) {
        super(props, state);
    }

    /**
     * Sets the validity state
     * @param val validity
     */
    public setValidityState(val: boolean) {
        if (microsoftTeams) {
            microsoftTeams.settings.setValidityState(val);
        }
    }

    /**
     * Returns the font size (default 16)
     */
    protected pageFontSize = (): number => {
        let sizeStr = window.getComputedStyle(document.getElementsByTagName("html")[0]).getPropertyValue("font-size");
        sizeStr = sizeStr.replace("px", "");
        let fontSize = parseInt(sizeStr, 10);
        if (!fontSize) {
            fontSize = 16;
        }
        return fontSize;
    }

    /**
     * Returns true of if hosted in Teams (in an Iframe)
     */
    protected inTeams = (): boolean => {
        try {
            return window.self !== window.top;
        } catch (e) {
            return true;
        }
    }

    /**
     * Updates the theme
     */
    protected updateTheme = (themeStr?: string): void => {
        let theme;
        switch (themeStr) {
            case "dark":
                theme = ThemeStyle.Dark;
                break;
            case "contrast":
                theme = ThemeStyle.HighContrast;
                break;
            case "default":
            default:
                theme = ThemeStyle.Light;
        }
        this.setState({ theme });
    }

    /**
     * Returns the value of a query variable
     */
    protected getQueryVariable = (variable: string): string | undefined => {
        const query = window.location.search.substring(1);
        const vars = query.split("&");
        for (const varPairs of vars) {
            const pair = varPairs.split("=");
            if (decodeURIComponent(pair[0]) === variable) {
                return decodeURIComponent(pair[1]);
            }
        }
        return undefined;
    }
}
