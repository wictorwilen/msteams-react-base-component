// Copyright (c) Wictor Wilén. All rights reserved.
// Licensed under the MIT license.

import * as React from "react";
import { render } from "react-dom";
import { themes } from "@fluentui/react-northstar";
import { ThemePrepared } from "@fluentui/styles";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * State interface for the Teams Base user interface React component
 */
export interface ITeamsBaseComponentState {
    /**
     * The Microsoft Teams theme style (Light, Dark, HighContrast)
     */
    theme: ThemePrepared<any>;
}

/**
 * Base implementation of the React based interface for the Microsoft Teams app
 */
export default class TeamsBaseComponent<P, S extends ITeamsBaseComponentState>
    extends React.Component<P, S> {

    /**
     * Static method to render the component
     * @param element DOM element to render the control in
     * @param props Properties
     */
    public static render<P>(element: HTMLElement, props: P) {
        return render(React.createElement(this, props), element);
    }

    /**
     * Returns true if hosted in Teams
     * @param timeout timeout in milliseconds, default = 1000
     * @returns a `Promise<boolean>`
     */
    protected inTeams = (timeout: number = 1000): Promise<boolean> => {
        return new Promise((resolve, reject) => {
            try {
                microsoftTeams.initialize(() => {
                    resolve(true);
                });
                setTimeout(() => {
                    resolve(false);
                }, timeout);
            } catch (e) {
                reject(e);
            }
        });
    }

    /**
     * Updates the theme
     */
    protected updateTheme = (themeStr?: string): void => {
        let theme: ThemePrepared<any>;
        switch (themeStr) {
            case "dark":
                theme = themes.teamsDark;
                break;
            case "contrast":
                theme = themes.teamsHighContrast;
                break;
            case "default":
            default:
                theme = themes.teams;
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
