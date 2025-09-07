export default {
  preset: "ts-jest",
  testEnvironment: "node",
  extensionsToTreatAsEsm: [".ts"],
  moduleNameMapper: {
    "^(\\.{1,2}/.*)\\.js$": "$1",
    chevrotain: "<rootDir>/node_modules/chevrotain/lib/chevrotain.mjs",
  },
  transform: {
    "^.+\\.ts$": [
      "ts-jest",
      {
        useESM: true,
        tsconfig: {
          module: "ES2022",
          target: "ES2022",
          esModuleInterop: true,
          allowSyntheticDefaultImports: true,
          moduleResolution: "node",
        },
      },
    ],
  },
  transformIgnorePatterns: ["node_modules/(?!(chevrotain)/)"],
  testMatch: ["**/__tests__/**/*.test.ts", "**/?(*.)+(spec|test).ts"],
  collectCoverageFrom: [
    "src/**/*.ts",
    "!src/**/*.d.ts",
    "!src/tests/**/*",
    "!src/cli/**/*",
  ],
  coverageDirectory: "coverage",
  coverageReporters: ["text", "lcov", "html"],
};
