// Copyright (c) Wictor Wilén. All rights reserved.
// Licensed under the MIT license.
// SPDX-License-Identifier: MIT

/**
 * @jest-environment jsdom
 */

// eslint-disable-next-line no-use-before-define
import * as useTeams from "./useTeams";

describe("checkInTeams", () => {

    let windowSpy: jest.SpyInstance;

    beforeEach(() => {
        jest.resetAllMocks();
        jest.clearAllMocks();
        windowSpy = jest.spyOn(window, "window", "get");
        jest.mock("@microsoft/teams-js", () => ({ app: {} }));
    });

    afterEach(() => {
        windowSpy.mockRestore();
    });

    it("Should return false if no Teams JS SDK", () => {
        jest.mock("@microsoft/teams-js", () => (undefined));
        expect(useTeams.checkInTeams()).toBeFalsy();
    });

    it("Should return false if Teams JS SDK and no frames or agents", () => {
        expect(useTeams.checkInTeams()).toBeFalsy();
    });

    it("Should return true if Teams JS SDK and correct agent", () => {
        windowSpy.mockImplementation(() => ({ navigator: { userAgent: "Something/Teams/Something" } }));
        expect(useTeams.checkInTeams()).toBeTruthy();
    });

    it("Should return true if Teams JS SDK and embedded page container", () => {
        windowSpy.mockImplementation(() => ({ name: "embedded-page-container", navigator: { userAgent: "" } }));
        expect(useTeams.checkInTeams()).toBeTruthy();
    });

    it("Should return true if Teams JS SDK and extension tab frame", () => {
        windowSpy.mockImplementation(() => ({ name: "extension-tab-frame", navigator: { userAgent: "" } }));
        expect(useTeams.checkInTeams()).toBeTruthy();
    });
});
