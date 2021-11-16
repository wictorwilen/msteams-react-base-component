// Copyright (c) Wictor Wilén. All rights reserved.
// Licensed under the MIT license.
// SPDX-License-Identifier: MIT

import { useEffect, useState } from "react";
import { unstable_batchedUpdates as batchedUpdates } from "react-dom";
import * as teamsJs from "@microsoft/teams-js";
import { app, pages } from "@microsoft/teams-js";
import { teamsDarkTheme, teamsHighContrastTheme, teamsTheme, ThemePrepared } from "@fluentui/react-northstar";

export const checkInTeams = (): boolean => {
    if (teamsJs === undefined) { // teams SDK JS not loaded
        return false;
    }
    if ((window.parent === window.self && (window as any).nativeInterface) ||
        window.navigator.userAgent.includes("Teams/") ||
        window.name === "embedded-page-container" ||
        window.name === "extension-tab-frame") {
        return true;
    }
    return false;
};

export const getQueryVariable = (variable: string): string | undefined => {
    const query = window.location.search.substring(1);
    const vars = query.split("&");
    for (const varPairs of vars) {
        const pair = varPairs.split("=");
        if (decodeURIComponent(pair[0]) === variable) {
            return decodeURIComponent(pair[1]);
        }
    }
    return undefined;
};

/**
 * Microsoft Teams React hook
 * @param options optional options
 * @returns A tuple with properties and methods
 * properties:
 *  - inTeams: boolean = true if inside Microsoft Teams
 *  - fullscreen: boolean = true if in full screen mode
 *  - theme: Fluent UI Theme
 *  - themeString: string - representation of the theme (default, dark or contrast)
 *  - context - the Microsoft Teams JS SDK context
 *  - host - quick access to the host properties
 * methods:
 *  - setTheme - manually set the theme
 */
export function useTeams(options?: { initialTheme?: string, setThemeHandler?: (theme?: string) => void }): [
    {
        inTeams?: boolean,
        fullScreen?: boolean,
        theme: ThemePrepared,
        themeString: string,
        context?: app.Context,
        host?: app.AppHostInfo
    }, {
        setTheme: (theme: string | undefined) => void
    }] {
    const [inTeams, setInTeams] = useState<boolean | undefined>(undefined);
    const [fullScreen, setFullScreen] = useState<boolean | undefined>(undefined);
    const [theme, setTheme] = useState<ThemePrepared>(teamsTheme);
    const [themeString, setThemeString] = useState<string>("default");
    const [initialTheme] = useState<string | undefined>((options && options.initialTheme) ? options.initialTheme : getQueryVariable("theme"));
    const [context, setContext] = useState<app.Context | undefined>(undefined);
    const [host, setHost] = useState<app.AppHostInfo | undefined>(undefined);

    const themeChangeHandler = (theme: string | undefined) => {
        setThemeString(theme || "default");
        switch (theme) {
            case "dark":
                setTheme(teamsDarkTheme);
                break;
            case "contrast":
                setTheme(teamsHighContrastTheme);
                break;
            case "default":
            default:
                setTheme(teamsTheme);
        }
    };

    const overrideThemeHandler = options?.setThemeHandler ? options.setThemeHandler : themeChangeHandler;

    useEffect(() => {
        // set initial theme based on options or query string
        if (initialTheme) {
            overrideThemeHandler(initialTheme);
        }

        app.initialize().then(() => {
            app.getContext().then(context => {
                batchedUpdates(() => {
                    setInTeams(true);
                    setContext(context);
                    setFullScreen(context.page.isFullScreen);
                    setHost(context.app.host);
                });
                overrideThemeHandler(context.app.theme);
                app.registerOnThemeChangeHandler(overrideThemeHandler);
                pages.registerFullScreenHandler((isFullScreen) => {
                    setFullScreen(isFullScreen);
                });
            }).catch(() => {
                setInTeams(false);
            });
        }).catch(() => {
            setInTeams(false);
        });

        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    return [{ inTeams, fullScreen, theme, context, themeString, host }, { setTheme: overrideThemeHandler }];
}
