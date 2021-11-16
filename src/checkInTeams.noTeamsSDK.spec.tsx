// Copyright (c) Wictor Wilén. All rights reserved.
// Licensed under the MIT license.
// SPDX-License-Identifier: MIT

/**
 * @jest-environment jsdom
 */

// eslint-disable-next-line no-use-before-define
import * as useTeams from "./useTeams";

jest.mock("@microsoft/teams-js", () => (undefined));

describe("checkInTeams", () => {

    it("Should return false if no Teams JS SDK", () => {
        expect(useTeams.checkInTeams()).toBeFalsy();
    });

});
