import * as exp from "./index";

describe("index", () => {
    it("Should export useTeams", () => {
        expect(exp.useTeams).toBeDefined();
    });

    it("Should export getQueryVariable", () => {
        expect(exp.getQueryVariable).toBeDefined();
    });

    it("Should export checkInTeams", () => {
        expect(exp.checkInTeams).toBeDefined();
    });
});
