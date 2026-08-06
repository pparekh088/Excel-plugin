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
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].[contenthash].js",
      clean: true,
    },
    resolve: { extensions: [".ts", ".tsx", ".js"] },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: { loader: "ts-loader", options: { compilerOptions: { noEmit: false } } },
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
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets/*", to: "assets/[name][ext]" },
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
