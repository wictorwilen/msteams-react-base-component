// Copyright (c) Wictor Wilén. All rights reserved.
// Licensed under the MIT license.
// SPDX-License-Identifier: MIT

import { useEffect, useState } from "react";
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

const getQueryVariable = (variable: string): string | undefined => {
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
 */
export function useTeams(options?: { initialTheme?: string }): [{ inTeams: boolean, fullScreen?: boolean, theme: ThemePrepared, context?: microsoftTeams.Context }, { setTheme: (theme: string | undefined) => void }] {
    const [inTeams, setInTeams] = useState<boolean>(false);
    const [fullScreen, setFullScreen] = useState<boolean | undefined>(undefined);
    const [theme, setTheme] = useState<ThemePrepared>(teamsTheme);
    const [initialTheme] = useState<string | undefined>((options && options.initialTheme) ? options.initialTheme : getQueryVariable("theme"));
    const [context, setContext] = useState<microsoftTeams.Context>();

    const themeChangeHandler = (theme: string | undefined) => {
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
    useEffect(() => {
        if (checkInTeams()) {
            setInTeams(true);
            if (inTeams) {
                microsoftTeams.initialize(() => {
                    microsoftTeams.getContext(context => {
                        setContext(context);
                        setFullScreen(context.isFullScreen);
                        themeChangeHandler(context.theme);
                    });
                    microsoftTeams.registerFullScreenHandler((isFullScreen) => {
                        setFullScreen(isFullScreen);
                    });
                    microsoftTeams.registerOnThemeChangeHandler(themeChangeHandler);
                });
            }
        }
    }, [inTeams]);

    useEffect(() => {
        themeChangeHandler(initialTheme);
    }, [initialTheme]);

    return [{ inTeams, fullScreen, theme, context }, { setTheme: themeChangeHandler }];
}
