// Copyright (c) Wictor Wilén. All rights reserved.
// Licensed under the MIT license.
// SPDX-License-Identifier: MIT

/**
 * @jest-environment jsdom
 */

// eslint-disable-next-line no-use-before-define
import React from "react";
import { render, waitFor } from "@testing-library/react";
import * as useTeams from "./useTeams";
import * as microsoftTeams from "@microsoft/teams-js";
import { Flex, Header, Provider } from "@fluentui/react-northstar";

jest.mock("@microsoft/teams-js");

describe("useTeams", () => {
    let spyCheckInTeams: jest.SpyInstance;
    let spyInitialize: jest.SpyInstance;
    let spyRegisterOnThemeChangeHandler: jest.SpyInstance;
    let spyRegisterFullScreenHandler: jest.SpyInstance;
    let spyGetContext: jest.SpyInstance;

    beforeEach(() => {
        jest.resetAllMocks();
        jest.clearAllMocks();
        spyCheckInTeams = jest.spyOn(useTeams, "checkInTeams");
        spyCheckInTeams.mockReturnValue(true);
        spyInitialize = jest.spyOn(microsoftTeams, "initialize");
        spyInitialize.mockImplementation(cb => {
            if (cb) { setTimeout(() => { cb(); }, 200); };
        });
        spyRegisterOnThemeChangeHandler = jest.spyOn(microsoftTeams, "registerOnThemeChangeHandler");
        spyRegisterFullScreenHandler = jest.spyOn(microsoftTeams, "registerFullScreenHandler");
        spyGetContext = jest.spyOn(microsoftTeams, "getContext");
        spyGetContext.mockImplementation((cb) => {
            // eslint-disable-next-line node/no-callback-literal
            cb({
                isFullScreen: false,
                theme: "default"
            });
        });
    });

    it("Should create the useTeams hook - empty by default", async () => {
        const App = () => {
            const [{ inTeams }] = useTeams.useTeams({});
            return (
                <div>{"" + inTeams}</div>
            );
        };

        spyCheckInTeams.mockReturnValue(false);

        const { container } = render(<App />);

        expect(spyInitialize).toBeCalledTimes(1);
        expect(spyCheckInTeams).toBeCalledTimes(1);
        expect(container.textContent).toBe("false");
    });

    it("Should create the useTeams hook - in teams", async () => {
        const App = () => {
            const [{ inTeams, themeString }] = useTeams.useTeams({});
            return (
                <div><div>{inTeams ? "true" : "false"}</div>,<div> {themeString}</div></div>
            );
        };

        const { container } = render(<App />);

        await waitFor(() => {
            expect(spyCheckInTeams).toBeCalledTimes(1);
            expect(spyInitialize).toBeCalledTimes(1);
            expect(spyGetContext).toBeCalledTimes(1);
            expect(spyRegisterFullScreenHandler).toBeCalledTimes(1);
            expect(spyRegisterOnThemeChangeHandler).toBeCalledTimes(1);
        });

        expect(container.textContent).toBe("true, default");
    });

    it("Should create the useTeams hook - not in teams", async () => {
        const App = () => {
            const [{ inTeams, themeString }] = useTeams.useTeams({});
            return (
                <div><div>{inTeams ? "true" : "false"}</div>,<div> {themeString}</div></div>
            );
        };

        spyCheckInTeams.mockReturnValue(false);

        const { container } = render(<App />);

        await waitFor(() => {
            expect(spyCheckInTeams).toBeCalledTimes(1);
            expect(spyInitialize).toBeCalledTimes(1);
        });

        expect(container.textContent).toBe("false, default");
    });

    it("Should create the useTeams hook with dark theme", async () => {
        const App = () => {
            const [{ inTeams, themeString }] = useTeams.useTeams({ initialTheme: "dark" });
            return (
                <div><div>{inTeams ? "true" : "false"}</div>,<div> {themeString}</div></div>
            );
        };

        spyGetContext.mockImplementation((cb) => {
            // eslint-disable-next-line node/no-callback-literal
            cb({
                isFullScreen: false,
                theme: "dark"
            });
        });

        const { container } = render(<App />);

        await waitFor(() => {
            expect(spyCheckInTeams).toBeCalledTimes(1);
            expect(spyInitialize).toBeCalledTimes(1);
            expect(spyGetContext).toBeCalledTimes(1);
        });

        expect(container.textContent).toBe("true, dark");
    });

    it("Should create the useTeams hook with contrast theme", async () => {
        const App = () => {
            const [{ inTeams, themeString }] = useTeams.useTeams({ initialTheme: "contrast" });
            return (
                <div><div>{inTeams ? "true" : "false"}</div>,<div> {themeString}</div></div>
            );
        };

        spyGetContext.mockImplementation((cb) => {
            // eslint-disable-next-line node/no-callback-literal
            cb({
                isFullScreen: false,
                theme: "contrast"
            });
        });

        const { container } = render(<App />);

        await waitFor(() => {
            expect(spyCheckInTeams).toBeCalledTimes(1);
            expect(spyInitialize).toBeCalledTimes(1);
            expect(spyGetContext).toBeCalledTimes(1);
        });

        expect(container.textContent).toBe("true, contrast");
    });

    it("Should create the useTeams hook with default theme, but switch to dark", async () => {
        const App = () => {
            const [{ inTeams, themeString }] = useTeams.useTeams({ initialTheme: "default" });
            return (
                <div><div>{inTeams ? "true" : "false"}</div>,<div> {themeString}</div></div>
            );
        };

        spyGetContext.mockImplementation((cb) => {
            // eslint-disable-next-line node/no-callback-literal
            cb({
                isFullScreen: false,
                theme: "dark"
            });
        });

        const { container } = render(<App />);

        await waitFor(() => {
            expect(spyCheckInTeams).toBeCalledTimes(1);
            expect(spyInitialize).toBeCalledTimes(1);
            expect(spyGetContext).toBeCalledTimes(1);
        });

        expect(container.textContent).toBe("true, dark");
    });

    it("Should create the useTeams hook with no theme, but switch to default", async () => {
        const App = () => {
            const [{ inTeams, themeString }] = useTeams.useTeams({});
            return (
                <div><div>{inTeams ? "true" : "false"}</div>,<div> {themeString}</div></div>
            );
        };

        const { container } = render(<App />);

        await waitFor(() => {
            expect(spyCheckInTeams).toBeCalledTimes(1);
            expect(spyInitialize).toBeCalledTimes(1);
            expect(spyGetContext).toBeCalledTimes(1);
        });

        expect(container.textContent).toBe("true, default");
    });

    it("Should call custom theme handler", async () => {
        const setThemeHandler = jest.fn();
        const App = () => {
            const [{ inTeams, themeString }] = useTeams.useTeams({ setThemeHandler });
            return (
                <div><div>{inTeams ? "true" : "false"}</div>,<div> {themeString}</div></div>
            );
        };

        const { container } = render(<App />);

        await waitFor(() => {
            expect(setThemeHandler).toBeCalledTimes(1);
        });

        expect(container.textContent).toBe("true, default");
    });

    it("Should not be fullscreen", async () => {
        const App = () => {
            const [{ fullScreen }] = useTeams.useTeams();
            return (
                <div><div>{fullScreen ? "true" : "false"}</div></div>
            );
        };

        const { container } = render(<App />);

        await waitFor(() => {
            expect(spyRegisterFullScreenHandler).toBeCalledTimes(1);
        });

        expect(container.textContent).toBe("false");
    });

    it("Should be fullscreen", async () => {
        const App = () => {
            const [{ fullScreen }] = useTeams.useTeams();
            return (
                <div><div>{fullScreen ? "true" : "false"}</div></div>
            );
        };

        spyGetContext.mockImplementation((cb) => {
            // eslint-disable-next-line node/no-callback-literal
            cb({
                isFullScreen: true,
                theme: "default"
            });
        });

        const { container } = render(<App />);

        await waitFor(() => {
            expect(spyRegisterFullScreenHandler).toBeCalledTimes(1);
        });

        expect(container.textContent).toBe("true");
    });

    it("Should call useEffect and render Fluent UI components", async () => {
        const HooksTab = () => {
            const [{ inTeams, theme }] = useTeams.useTeams({});
            const [message, setMessage] = React.useState("Loading...");

            React.useEffect(() => {
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

        const { container } = render(<HooksTab />);

        await waitFor(() => {
            expect(container.textContent).toBe("In Microsoft Teams!");
        });
    });

    it("Should run the functional component three times", async () => {
        const ping = jest.fn();
        const pingEffect = jest.fn();

        const spyAppInit = jest.spyOn(microsoftTeams.appInitialization, "notifyAppLoaded");
        spyAppInit.mockImplementation(jest.fn());

        const HooksTab = () => {
            const [{ inTeams }] = useTeams.useTeams();

            React.useEffect(() => {
                pingEffect();
                if (inTeams) {
                    microsoftTeams.appInitialization.notifyAppLoaded();
                }
            }, [inTeams]);

            ping();

            return (
                <h1>Test</h1>
            );
        };

        render(<HooksTab />);

        await waitFor(() => {
            expect(ping).toBeCalledTimes(3);
            expect(pingEffect).toBeCalledTimes(2);
        });

    });
});
