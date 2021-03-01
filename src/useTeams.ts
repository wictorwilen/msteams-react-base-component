// Copyright (c) Wictor Wilén. All rights reserved.
// Licensed under the MIT license.
// SPDX-License-Identifier: MIT

import { useEffect, useState } from "react";
import { unstable_batchedUpdates as batchedUpdates } from "react-dom";
import * as microsoftTeams from "@microsoft/teams-js";
import { teamsDarkTheme, teamsHighContrastTheme, teamsTheme, ThemePrepared } from "@fluentui/react-northstar";

export const checkInTeams = (): boolean => {
    // eslint-disable-next-line dot-notation
    const microsoftTeamsLib = microsoftTeams || window["microsoftTeams"];

    if (!microsoftTeamsLib) {
        return false; // the Microsoft Teams library is for some reason not loaded
    }

    if ((window.parent === window.self && (window as any).nativeInterface) ||
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
 * methods:
 *  - setTheme - manually set the theme
 */
export function useTeams(options?: { initialTheme?: string, setThemeHandler?: (theme?: string) => void }): [
    {
        inTeams?: boolean,
        fullScreen?: boolean,
        theme: ThemePrepared,
        themeString: string,
        context?: microsoftTeams.Context
    }, {
        setTheme: (theme: string | undefined) => void
    }] {
    const [inTeams, setInTeams] = useState<boolean | undefined>(undefined);
    const [fullScreen, setFullScreen] = useState<boolean | undefined>(undefined);
    const [theme, setTheme] = useState<ThemePrepared>(teamsTheme);
    const [themeString, setThemeString] = useState<string>("default");
    const [initialTheme] = useState<string | undefined>((options && options.initialTheme) ? options.initialTheme : getQueryVariable("theme"));
    const [context, setContext] = useState<microsoftTeams.Context>();

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
        const isInTeams = checkInTeams();
        if (isInTeams) {
            microsoftTeams.initialize(() => {
                microsoftTeams.getContext(context => {
                    batchedUpdates(() => {
                        setInTeams(true);
                        setContext(context);
                        setFullScreen(context.isFullScreen);
                    });
                    overrideThemeHandler(context.theme);
                });
                microsoftTeams.registerFullScreenHandler((isFullScreen) => {
                    setFullScreen(isFullScreen);
                });
                microsoftTeams.registerOnThemeChangeHandler(overrideThemeHandler);
            });
        } else {
            setInTeams(false);
            microsoftTeams.initialize();
        }
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    return [{ inTeams, fullScreen, theme, context, themeString }, { setTheme: overrideThemeHandler }];
}
