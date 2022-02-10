# Microsoft Teams UI Controls base component

[![npm version](https://badge.fury.io/js/msteams-react-base-component.svg)](https://www.npmjs.com/package/msteams-react-base-component)
[![npm](https://img.shields.io/npm/dt/msteams-react-base-component.svg)](https://www.npmjs.com/package/msteams-react-base-component)
[![MIT](https://img.shields.io/npm/l/msteams-react-base-component.svg)](https://github.com/wictorwilen/msteams-react-base-component/blob/master/LICENSE)
[![GitHub issues](https://img.shields.io/github/issues/wictorwilen/msteams-react-base-component.svg)](https://github.com/wictorwilen/msteams-react-base-component/issues)
[![GitHub closed issues](https://img.shields.io/github/issues-closed/wictorwilen/msteams-react-base-component.svg)](https://github.com/wictorwilen/msteams-react-base-component/issues?q=is%3Aissue+is%3Aclosed)

This is a set of React hooks and providers based on the Microsoft Teams JavaScript SDK, the Fluent UI components and Microsoft Graph Toolkit, which is used when generating Microsoft Teams Apps using the [Microsoft Teams Yeoman Generator](https://aka.ms/yoteams).

 | @master | @preview |
 :--------:|:---------:
 ![Build Status](https://img.shields.io/github/workflow/status/wictorwilen/msteams-react-base-component/msteams-react-base-component%20CI/master)|![Build Status](https://img.shields.io/github/workflow/status/wictorwilen/msteams-react-base-component/msteams-react-base-component%20CI/preview)

# Usage

## `useTeams` hook

To use this package in a Teams tab or extension import the `useTeams` Hook and then call it inside a functional component.

``` TypeScript
const [{inTeams}] = useTeams();
```

The `useTeams` hook will return a tuple of where an object of properties are in the first field and an object of methods in the second.

> **NOTE**: using the hook will automatically call `microsoftTeams.initialize()` and `microsoftTeams.getContext()` if the Microsoft Teams JS SDK is available.

### useTeams Hook arguments

The `useTeams` hook can take an *optional* object argument:

| Argument | Description |
|----------|-------------|
| `initialTheme?: string` | Manually set the initial theme (`default`, `dark` or `contrast`) |
| `setThemeHandler?: (theme?: string) => void` | Custom handler for themes |

### Available properties

| Property name | Type | Description |
|---------------|------|-------------|
| `inTeams` | boolean? | `true` if hosted in Teams and `false` for outside of Microsoft Teams |
| `fullScreen` | boolean? | `true` if the Tab is in full-screen, otherwise `false` |
| `themeString` | string | The value of `default`, `dark` or `contrast` |
| `theme` | ThemePrepared | The Fluent UI Theme object for the current theme |
| `context` | `microsoftTeams.Context?` | `undefined` while the Tab is loading or if not hosted in Teams, set to a value once the Tab is initialized and context available |

### Available methods

| Method name | Description |
|-------------|-------------|
| `setTheme(theme?: string)` | Method for manually setting the theme |

## Full example

Example of usage:

```  TypeScript
import * as React from "react";
import { Provider, Flex, Header } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";

/**
 * Implementation of the hooks Tab content page
 */
export const HooksTab = () => {
    const [{ inTeams, theme }] = useTeams({});
    const [message, setMessage] = useState("Loading...");

    useEffect(() => {
        if (inTeams === true) {
            setMessage("In Microsoft Teams!");
        } else {
            if (inTeams !== undefined) {
                setMessage("Not in Microsoft Teams");
            }
        }
    }, [inTeams]);

    return (
        <Provider theme={theme}>
            <Flex fill={true}>
                <Flex.Item>
                    <Header content={message} />
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
```

### Additional helper methods

The package also exports two helper methods, both used internally by the `useTeams` hook.

`getQueryVariable(name: string): string` - returns the value of the query string variable identified by the name.

`checkInTeams(): boolean` - returns true if hosted inside Microsoft Teams.

## The `TeamsSsoProvider`

The `TeamsSsoProvider` allows for a minimal coding experience when using Teams SSO tabs and authorization tokens. It works with Microsoft Teams SSO tabs that are opened inside of Microsoft Teams, as well as outside of Microsoft Teams.

## Importing the provider

Import he provider in your React project as follows:

``` TypeScript
import 
  { TeamsSsoProvider, useTeamsSsoContext, TeamsSsoContext } 
  from "msteams-react-base-component/lib/esm/TeamsSsoProvider";
```

### Configuration

The `TeamsSsoProvider` requires the following configuration properties:

| Property | Description |
|-|-|
| `appIdUri` | *Required* The Azure AD App ID URI |
| `appId` | *Optional* The Azure AD App ID |
| `scopes` | *Optional* The Scopes to use. Default to `[""]` |
| `redirectUri` | *Optional* The Redirect URI to be used for the AAD App |
| `useMgt` | *Optional* Boolean value indicating if the Microsoft Graph Toolkit auth provider should be initialized |
| `autoLogin` | *Optional* Boolean value indicating if the provider should automatically log in the user. Defaults to true. If set to false, then `login` of the `TeamsSsoProvider` context object has to be called. |

### Context

The `TeamsSsoProvider` contains the following context variables and methods

| Variable/method | Description  |
|-|-|
| `token` | The SSO/access token. `undefined` if not set |
| `name` | The user name. Defaults to `undefined` |
| `error` | Any error message. `undefined` if no errors |
| `logout()` | Signs the user out of the application |
| `login()` | Signs the user in to the application |

### Usage with the `TeamsSsoContext.Consumer`

To make an SSO token available to all your components, in the tree below the provider, use the following approach:

``` TypeScript
export const App = () => {
    return (
        <Provider theme={theme}>
            <TeamsSsoProvider
              appIdUri={process.env.TAB_APP_URI as string}>
              <TeamsSsoContext.Consumer>
              { state => (
                <div>Your token is <b>${state.token}</b></div> 
              )}
              </TeamsSsoContext.Consumer>
            </TeamsSsoProvider>
        </Provider>
    );
}
```

### Usage with the `useTeamsSsoContext` hook

To access the token in code and sub components use the following method:

``` TypeScript

const MyComponent = () => {
    const { token } = useTeamsSsoContext();
    return <div>Your token is <b>${token}</b></div>;
}

export const App = () => {
    return (
        <Provider theme={theme}>
            <TeamsSsoProvider
              appIdUri={process.env.TAB_APP_URI as string}>
                <MyComponent />
            </TeamsSsoProvider>
        </Provider>
    );
}
```

### Usage outside of Teams

To support usage outside of Teams (for instance to use as a PWA), you also need to specify the `appId` and `redirectUri` properties:

``` TypeScript
export const App = () => {
    return (
        <Provider theme={theme}>
            <TeamsSsoProvider
              appIdUri={process.env.TAB_APP_URI as string}
              appId={process.env.TAB_APP_ID as string}
              redirectUri={process.env.TAB_APP_REDIRECT as string}>
              <TeamsSsoContext.Consumer>
              { state => (
                <div>Your token is <b>${state.token}</b></div> 
              )}
              </TeamsSsoContext.Consumer>
            </TeamsSsoProvider>
        </Provider>
    );
}
```

### Scopes

To request more scopes on the access token, than an empty SSO token, you need to specify the scopes in the `scopes` property. Note, that requesting scopes will popup a request for consent dialog, even when inside Microsoft Teams.

``` TypeScript
export const App = () => {
    return (
        <Provider theme={theme}>
            <TeamsSsoProvider
              appIdUri={process.env.TAB_APP_URI as string}
              appId={process.env.TAB_APP_ID as string}
              redirectUri={process.env.TAB_APP_REDIRECT as string}
              scopes={["Presence.Read", "User.Read"]}>
              <TeamsSsoContext.Consumer>
              { state => (
                <div>Your token is <b>${state.token}</b></div> 
              )}
              </TeamsSsoContext.Consumer>
            </TeamsSsoProvider>
        </Provider>
    );
}
```

### Usage with Microsoft Graph Toolkit

For seamless use with the Microsoft Graph Toolkit (MGT), the only thing you need to specify is the `useMgt` property and set it to true:

``` TypeScript
export const App = () => {
    return (
        <Provider theme={theme}>
            <TeamsSsoProvider
              appIdUri={process.env.TAB_APP_URI as string}
              appId={process.env.TAB_APP_ID as string}
              redirectUri={process.env.TAB_APP_REDIRECT as string}
              scopes={["Presence.Read", "User.Read"]}
              useMgt={true}>
              <Person personQuery="me" showPresence={true} />
            </TeamsSsoProvider>
        </Provider>
    );
}
```
# License

Copyright (c) Wictor Wilén. All rights reserved.

Licensed under the MIT license.
