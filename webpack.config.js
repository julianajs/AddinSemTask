const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = {
  entry: {
    commands: "./src/commands/commands.js"
  },
  output: {
    clean: true,
    filename: "[name].js",
    path: path.resolve(__dirname, "dist")
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: "./src/commands/commands.html",
      chunks: ["commands"]
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: "src/assets", to: "assets" },
        { from: "manifest.xml", to: "manifest.xml" }
      ]
    })
  ],
  devServer: {
    headers: {
      "Access-Control-Allow-Origin": "*"
    },
    server: "https",
    port: 3000
  }
};
