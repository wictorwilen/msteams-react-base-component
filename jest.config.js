module.exports = {
    name: "client",
    displayName: "client",
    rootDir: "./",
    globals: {
        "ts-jest": {
            tsconfig: "<rootDir>/tsconfig.json",
            diagnostics: {
                ignoreCodes: [151001]
            }
        }
    },
    preset: "ts-jest/presets/js-with-ts",
    testMatch: [
        "<rootDir>/**/*.spec.(ts|tsx|js)"
    ],
    collectCoverageFrom: [
        "/**/*.{js,jsx,ts,tsx}",
        "!<rootDir>/node_modules/"
    ],
    coverageReporters: [
        "text", "html"
    ]
};
