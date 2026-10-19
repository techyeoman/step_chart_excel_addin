/* Webpack build for the Probability Step Chart add-in.
 *
 * Development: files are served from https://localhost:3000 using the Office
 * dev certificates (run `npm run certs` once).
 * Production:  set ADDIN_PUBLIC_URL to the HTTPS origin/path that will host
 * the built files, e.g. ADDIN_PUBLIC_URL=https://addins.example.com/stepchart/
 * The manifests copied into dist/ are rewritten to point at that URL.
 */
const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = process.env.ADDIN_PUBLIC_URL || "https://localhost:3000/";

async function getHttpsOptions() {
  const devCerts = require("office-addin-dev-certs");
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

/** `npm run preview`: the task pane against an in-memory workbook, in any browser. */
function previewConfig() {
  return {
    mode: "development",
    devtool: "source-map",
    entry: { preview: "./src/preview/preview.tsx" },
    output: { path: path.resolve(__dirname, "dist-preview"), filename: "[name].js", clean: true },
    resolve: { extensions: [".ts", ".tsx", ".js"] },
    module: {
      rules: [{ test: /\.tsx?$/, exclude: /node_modules/, use: { loader: "ts-loader", options: { transpileOnly: true } } }],
    },
    performance: { hints: false },
    plugins: [new HtmlWebpackPlugin({ filename: "index.html", template: "./src/preview/preview.html" })],
    devServer: { port: 8123, server: "http", static: false },
  };
}

module.exports = async (env, options) => {
  if (env.preview) return previewConfig();
  const dev = options.mode === "development";
  const prodUrl = urlProd.endsWith("/") ? urlProd : urlProd + "/";
  const rewriteUrls = (content) =>
    dev ? content : content.toString().replace(new RegExp(urlDev, "g"), prodUrl);

  return {
    devtool: dev ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/index.tsx",
      commands: "./src/commands/commands.ts",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: { loader: "ts-loader", options: { transpileOnly: true } },
        },
      ],
    },
    performance: { hints: false },
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
          { from: "assets/*.png", to: "assets/[name][ext]" },
          { from: "src/static/*.html", to: "[name][ext]", transform: rewriteUrls },
          // The about page doubles as the site's landing page (GitHub Pages root).
          { from: "src/static/about.html", to: "index.html", transform: rewriteUrls },
          { from: "manifest.json", to: "manifest.json", transform: rewriteUrls },
          { from: "manifest.xml", to: "manifest.xml", transform: rewriteUrls },
        ],
      }),
    ],
    devServer: {
      headers: { "Access-Control-Allow-Origin": "*" },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };
};
