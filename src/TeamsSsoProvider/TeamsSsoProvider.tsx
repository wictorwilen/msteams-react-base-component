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
import { TeamsSsoContext, TeamsSsoStatus } from "./TeamsSsoContext";

/**
 * Teams React SSO Provider
 * @param props properties
 * @returns The TeamsSsoProvider
 */
export const TeamsSsoProvider = (props: React.PropsWithChildren<TeamsSsoProviderProps>) => {
    const [token, setToken] = React.useState<string | undefined>(undefined);
    const [name, setName] = React.useState<string | undefined>(undefined);
    const [error, setError] = React.useState<string | undefined>(undefined);
    const [isLoggingIn, setIsLoggingIn] = React.useState<boolean>(false);
    const [status, setStatus] = React.useState<TeamsSsoStatus>("Unknown");
    const [mgtLoaded, setMgtLoaded] = React.useState<boolean>(false);

    const [{ inTeams }] = useTeams();

    const login = React.useCallback((): void => {
        if (inTeams === true) {
            setStatus("LoggingIn");
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
                            microsoftTeams.appInitialization.notifySuccess();
                        }).catch(err => {
                            if (err instanceof msal.InteractionRequiredAuthError || err instanceof msal.BrowserAuthError) {
                                if (props.scopes && props.redirectUri && props.appId) {
                                    microsoftTeams.authentication.authenticate({
                                        successCallback: (result) => {
                                            console.log(result);
                                        },
                                        failureCallback: (reason) => {
                                            console.error(reason);
                                        },
                                        url: props.redirectUri
                                    });
                                    // TODO: Teams popup - needs a handler page!
                                    // msalInstance.loginPopup({
                                    //     scopes: props.scopes,
                                    //     redirectUri: props.redirectUri
                                    // }).then(result => {
                                    //     setToken(result.accessToken);
                                    //     const account = msalInstance.getAllAccounts();
                                    //     if (account.length === 1) {
                                    //         setName(account[0].name);
                                    //     }
                                    //     setError(undefined);
                                    //     microsoftTeams.appInitialization.notifySuccess();
                                    // }).catch(err => {
                                    //     setError(err);
                                    // });
                                }
                            } else {
                                setError(err);
                            }
                        });
                    } else {
                        if (props.scopes && !props.appId) {
                            const message = "When scopes is specified an appId is also required";
                            setError(message);
                            microsoftTeams.appInitialization.notifyFailure({
                                reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                                message
                            });
                        }
                        setToken(token);
                        microsoftTeams.appInitialization.notifySuccess();
                    }
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
                const msalConfig: msal.Configuration = {
                    auth: {
                        clientId: props.appId,
                        redirectUri: props.redirectUri
                    }
                };
                const msalInstance = new msal.PublicClientApplication(msalConfig);
                setStatus("LoggingIn");
                msalInstance.ssoSilent({ scopes: props.scopes }).then(result => {
                    setToken(result.accessToken);
                    const account = msalInstance.getAllAccounts();
                    if (account.length === 1) {
                        setName(account[0].name);
                    }
                    setError(undefined);
                    setStatus("LoggedIn");
                }).catch(err => {
                    if (err instanceof msal.InteractionRequiredAuthError || err instanceof msal.BrowserAuthError) {
                        setStatus("WaitingForUser");
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
                                setStatus("LoggedIn");
                                setError(undefined);
                            }).catch(err => {
                                setStatus("Error");
                                setError(err);
                            });
                        } else {
                            throw new Error("Missing redirectUri");
                        }
                    } else {
                        setStatus("Error");
                        setError(err);
                    }
                });

            } else {
                setStatus("Error");
                throw new Error("Missing scopes and/or appId");
            }
        }
    }, [inTeams, props.appId, props.appIdUri, props.scopes, props.redirectUri]);

    useEffect(() => {
        if (inTeams !== undefined) {
            if (!isLoggingIn) {
                if (props.autoLogin === true) {
                    console.log("auto logging in");
                    setIsLoggingIn(true);
                    login();
                }
            }
        }
    }, [login, inTeams, isLoggingIn, props.autoLogin]);

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
                setMgtLoaded(true);
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
        } else {
            console.warn("Nothing to logout from, missing appId");
        }
    }, [props.appId]);

    const memoedToken = React.useMemo(() => ({ token, name, error, logout, login, status, mgtLoaded }), [token, name, error, logout, login, status, mgtLoaded]);

    return (<TeamsSsoContext.Provider value={memoedToken} >
        {props.children}
    </TeamsSsoContext.Provider >
    );
};
