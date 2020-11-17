# Microsoft Teams UI Controls base component

[![npm version](https://badge.fury.io/js/msteams-react-base-component.svg)](https://www.npmjs.com/package/msteams-react-base-component)
[![npm](https://img.shields.io/npm/dt/msteams-react-base-component.svg)](https://www.npmjs.com/package/msteams-react-base-component)
[![MIT](https://img.shields.io/npm/l/msteams-react-base-component.svg)](https://github.com/wictorwilen/msteams-react-base-component/blob/master/LICENSE)
[![GitHub issues](https://img.shields.io/github/issues/wictorwilen/msteams-react-base-component.svg)](https://github.com/wictorwilen/msteams-react-base-component/issues)
[![GitHub closed issues](https://img.shields.io/github/issues-closed/wictorwilen/msteams-react-base-component.svg)](https://github.com/wictorwilen/msteams-react-base-component/issues?q=is%3Aissue+is%3Aclosed)

This is a React hook based on the Microsoft Teams JavaScript SDK and the Fluent UI components, which is used when generating Microsoft Teams Apps using the [Microsoft Teams Yeoman Generator](https://aka.ms/yoteams).

 | @master | @preview |
 :--------:|:---------:
 ![Build Status](https://img.shields.io/github/workflow/status/wictorwilen/msteams-react-base-component/msteams-react-base-component%20CI/master)|![Build Status](https://img.shields.io/github/workflow/status/wictorwilen/msteams-react-base-component/msteams-react-base-component%20CI/preview)

## Usage

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

## License

Copyright (c) Wictor Wilén. All rights reserved.

Licensed under the MIT license.
