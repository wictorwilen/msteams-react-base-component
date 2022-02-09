// Copyright (c) Wictor Wilén. All rights reserved.
// Licensed under the MIT license.
// SPDX-License-Identifier: MIT
import * as React from "react";
import { useEffect } from "react";
import { useTeams } from "../useTeams";
import "isomorphic-fetch";
import { TeamsSsoProviderProps } from "./TeamsSsoProviderProps";
import jwtDecode from "jwt-decode";
import * as msal from "@azure/msal-browser";
import * as microsoftTeams from "@microsoft/teams-js";
import { Providers } from "@microsoft/mgt-element/dist/es6/providers/Providers";
import { ProviderState } from "@microsoft/mgt-element/dist/es6/providers/IProvider";
import { SimpleProvider } from "@microsoft/mgt-element/dist/es6/providers/SimpleProvider";
import { TeamsSsoContext } from "./TeamsSsoContext";

/**
 * Teams React SSO Provider
 * @param props properties
 * @returns The TeamsSsoProvider
 */
export const TeamsSsoProvider = (props: React.PropsWithChildren<TeamsSsoProviderProps>) => {
    const [token, setToken] = React.useState<string | undefined>(undefined);
    const [name, setName] = React.useState<string | undefined>(undefined);
    const [error, setError] = React.useState<string | undefined>(undefined);

    const [{ inTeams }] = useTeams();

    useEffect(() => {
        if (inTeams === true) {

            microsoftTeams.authentication.getAuthToken({
                successCallback: (token: string) => {
                    const decoded: { [key: string]: any; } = jwtDecode(token) as { [key: string]: any; };
                    setError(undefined);
                    setName(decoded!.name);
                    if (props.scopes && props.appId) {
                        // if we have scopes, then we need to use MSAL
                        const msalConfig = {
                            auth: {
                                clientId: props.appId
                            }
                        };
                        const msalInstance = new msal.PublicClientApplication(msalConfig);

                        msalInstance.ssoSilent({
                            loginHint: decoded.upn,
                            scopes: props.scopes
                        }).then(result => {
                            setToken(result.accessToken);
                            const account = msalInstance.getAllAccounts();
                            if (account.length === 1) {
                                setName(account[0].name);
                            }
                            setError(undefined);
                        }).catch(err => {
                            if (err instanceof msal.InteractionRequiredAuthError) {
                                if (props.scopes && props.redirectUri && props.appId) {
                                    msalInstance.loginPopup({
                                        scopes: props.scopes,
                                        redirectUri: props.redirectUri
                                    }).then(result => {
                                        setToken(result.accessToken);
                                        const account = msalInstance.getAllAccounts();
                                        if (account.length === 1) {
                                            setName(account[0].name);
                                        }
                                        setError(undefined);
                                    }).catch(err => {
                                        setError(err);
                                    });
                                }
                            } else {
                                setError(err);
                            }
                        });
                    } else {
                        setToken(token);
                    }
                    microsoftTeams.appInitialization.notifySuccess();
                },
                failureCallback: (message: string) => {
                    setError(message);
                    microsoftTeams.appInitialization.notifyFailure({
                        reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                        message
                    });
                },
                resources: [props.appIdUri]
            });

        } else if (inTeams === false) {
            if (props.scopes && props.appId) {
                const msalConfig = {
                    auth: {
                        clientId: props.appId
                    }
                };
                const msalInstance = new msal.PublicClientApplication(msalConfig);

                msalInstance.ssoSilent({ scopes: props.scopes }).then(result => {
                    setToken(result.accessToken);
                    const account = msalInstance.getAllAccounts();
                    if (account.length === 1) {
                        setName(account[0].name);
                    }
                    setError(undefined);
                }).catch(err => {
                    if (err instanceof msal.InteractionRequiredAuthError) {
                        if (props.scopes && props.redirectUri && props.appId) {
                            msalInstance.loginPopup({
                                scopes: props.scopes,
                                redirectUri: props.redirectUri
                            }).then(result => {
                                setToken(result.accessToken);
                                const account = msalInstance.getAllAccounts();
                                if (account.length === 1) {
                                    setName(account[0].name);
                                }
                                setError(undefined);
                            }).catch(err => {
                                setError(err);
                            });
                        }
                    } else {
                        setError(err);
                    }
                });

            }
        }
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [inTeams]);

    useEffect(() => {
        if (props.useMgt) {
            Providers.globalProvider = new SimpleProvider(() => {
                if (token) {
                    return Promise.resolve(token);
                } else {
                    return Promise.reject(new Error("No token"));
                }
            });

            if (token) {
                Providers.globalProvider.setState(ProviderState.SignedIn);
            } else {
                Providers.globalProvider.setState(ProviderState.Loading);
            }
        }

    }, [token, props.useMgt]);

    const logout = React.useCallback((): void => {
        if (props.appId) {
            const msalConfig = {
                auth: {
                    clientId: props.appId
                }
            };
            const msalInstance = new msal.PublicClientApplication(msalConfig);
            setToken(undefined);
            setError(undefined);
            setName(undefined);
            msalInstance.logoutRedirect();
        }
    }, [props.appId]);

    const memoedToken = React.useMemo(() => ({ token, name, error, logout }), [token, name, error, logout]);

    return (<TeamsSsoContext.Provider value={memoedToken} >
        {props.children}
    </TeamsSsoContext.Provider >
    );
};
