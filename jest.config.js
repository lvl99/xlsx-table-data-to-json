module.exports = {
  preset: "ts-jest",
  testEnvironment: "node",
  collectCoverageFrom: ["lib/**/*.ts"],
  coverageReporters: ["text-summary", "html"],
  watchPathIgnorePatterns: ["node_modules/"],
};
