/* eslint-disable no-undef */
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://ledger-addin.example.com/"; // TODO: real hosting URL at pilot time

module.exports = async (env, options) => {
  const dev = options.mode === "development";

  /** @type {import("webpack").Configuration & {devServer?: object}} */
  const config = {
    devtool: dev ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/index.tsx",
      commands: "./src/commands/commands.ts",
      // Custom functions run in their own JS runtime (handoff §8).
      functions: "./src/functions/functions.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      // The custom-functions bundle must keep a stable name: the manifest
      // references it by URL and cannot follow a content hash.
      filename: (pathData) =>
        pathData.chunk?.name === "functions" ? "functions.js" : "[name].[contenthash].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js"],
      // The engine is consumed from source so the add-in and the eval harness
      // always run byte-identical audit logic (see DECISIONS D-011).
      alias: { "ledger-engine": path.resolve(__dirname, "../engine/src/index.ts") },
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: {
            loader: "ts-loader",
            options: {
              compilerOptions: { noEmit: false },
              // Engine sources live outside this package's rootDir.
              transpileOnly: false,
              onlyCompileBundledFiles: true,
            },
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["functions"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets/*", to: "assets/[name][ext]" },
          { from: "src/functions/functions.json", to: "functions.json" },
          {
            from: "manifest.xml",
            to: "manifest.xml",
            transform(content) {
              return dev
                ? content
                : content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            },
          },
        ],
      }),
    ],
  };

  if (dev) {
    // Lazy-require so production builds don't need the dev-certs package to
    // have generated certificates yet.
    const devCerts = require("office-addin-dev-certs");
    config.devServer = {
      hot: true,
      headers: { "Access-Control-Allow-Origin": "*" },
      server: { type: "https", options: await devCerts.getHttpsServerOptions() },
      port: 3000,
      static: { directory: path.join(__dirname, "dist") },
    };
  }

  return config;
};
