// Copyright (c) Wictor Wilén. All rights reserved.
// Licensed under the MIT license.
// SPDX-License-Identifier: MIT

/**
 * TeamsSsoProvider settings
 */
export interface TeamsSsoProviderProps {
    /**
     * Application ID
     */
    appId?: string;
    /**
     * Application ID URI
     */
    appIdUri: string;
    /**
     * Scopes. Defaults to empty scope
     */
    scopes?: string[];
    /**
     * Redirect Uri
     */
    redirectUri?: string;
    /**
     * Set to true to initialize Microsoft Graph Toolkit authorization provider
     */
    useMgt?: boolean;
}
