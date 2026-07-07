const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

// Webpack supports async config functions — used here to await the trusted dev cert options.
module.exports = async (env, options) => {
  const devMode = options.mode !== "production";
  const httpsOptions = devMode
    ? await require("office-addin-dev-certs").getHttpsServerOptions()
    : {};

  return {
    entry: {
      taskpane: "./src/taskpane/taskpane.js",
      "launchevent-v10": "./src/launchevent/launchevent.js"
    },
    output: {
      filename: "[name].js",
      path: path.resolve(__dirname, "dist"),
      clean: true
    },
    plugins: [
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"]
      }),
      new HtmlWebpackPlugin({
        filename: "commands-v10.html",
        template: "./src/commands/commands.html",
        chunks: ["launchevent-v10"],
        scriptLoading: "blocking",
        hash: true
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "manifest.xml",
            to: "manifest.xml"
          }
        ]
      })
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*"
      },
      server: {
        type: "https",
        options: httpsOptions
      },
      port: 3000,
      hot: false
    }
  };
};
