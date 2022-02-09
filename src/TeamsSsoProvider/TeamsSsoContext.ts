// Copyright (c) Wictor Wilén. All rights reserved.
// Licensed under the MIT license.
// SPDX-License-Identifier: MIT

import React from "react";

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
};

export const TeamsSsoContext = React.createContext<Partial<TeamsSsoContextProps>>({

});
