// Copyright (c) Wictor Wilén. All rights reserved.
// Licensed under the MIT license.
// SPDX-License-Identifier: MIT

import React from "react";

export declare type TeamsSsoStatus = "Unknown" | "LoggingIn" | "LoggedIn" | "WaitingForUser" | "Error";

/**
 * The Teams SSO Context
 */
export interface TeamsSsoContextProps {
    /**
     * The token
     */
    token: string;
    /**
     * User name
     */
    name: string;
    /**
     * Optional error message
     */
    error?: string;

    /**
     * Logout method
     */
    logout: () => void;

    /**
     * Login method
     */
    login: () => void;

    status: string;

    mgtLoaded: boolean;
};

export const TeamsSsoContext = React.createContext<Partial<TeamsSsoContextProps>>({

});
