/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");
const fs = require("fs");
const { execSync } = require("child_process");
const webpack = require("webpack");

const urlDev = "https://localhost:3000/";
const urlProd = "https://cdn-msoffice.glassnode.com/msexcel/";

/* global require, module, process, __dirname */

// Writes version.json next to the build: the add-in version, the commit it was built from and
// whether the working tree had uncommitted changes. Served with the add-in, it records what
// the CDN runs (DEVELOPMENT.md -> Releasing).
class VersionFilePlugin {
  apply(compiler) {
    compiler.hooks.thisCompilation.tap("VersionFilePlugin", (compilation) => {
      compilation.hooks.processAssets.tap(
        { name: "VersionFilePlugin", stage: webpack.Compilation.PROCESS_ASSETS_STAGE_ADDITIONAL },
        () => {
          const versionSource = fs.readFileSync(path.join(__dirname, "src/version.ts"), "utf8");
          const version = (versionSource.match(/ADDIN_VERSION = '([^']+)'/) || [])[1] || "unknown";
          const git = (cmd) => {
            try {
              return execSync(cmd, { cwd: __dirname, stdio: ["ignore", "pipe", "ignore"] }).toString().trim();
            } catch (e) {
              return null;
            }
          };
          const commit = git("git rev-parse HEAD");
          const status = git("git status --porcelain");
          const info = {
            version,
            commit: commit || "unknown",
            tag: git("git describe --tags --exact-match HEAD") || null,
            dirty: status === null ? null : status.length > 0,
          };
          compilation.emitAsset("version.json", new webpack.sources.RawSource(JSON.stringify(info, null, 2) + "\n"));
        }
      );
    });
  }
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
      functions: "./src/functions/functions.ts",
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: [/node_modules/, /\.test\.ts$/, /__tests__/],
          use: {
            loader: "babel-loader"
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new VersionFilePlugin(),
      new CustomFunctionsMetadataPlugin({
        output: "functions.json",
        input: "./src/functions/functions.ts",
      }),
      new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"],
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev + "(?:public/)?", "g"), urlProd);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
    ],
    watchOptions: {
      ignored: ['**/dist'],
      aggregateTimeout: 300,
      poll: false
    },
    devServer: {
      static: {
        directory: path.join(__dirname, "dist"),
        publicPath: "/public",
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      proxy: [
        {
          context: ['/api/glassnode'],
          target: 'https://api.glassnode.com',
          changeOrigin: true,
          secure: true,
          pathRewrite: {
            '^/api/glassnode': ''
          },
          logLevel: 'debug',
          onProxyReq: function(proxyReq, req, res) {
            console.log('Proxying request to:', proxyReq.protocol + '//' + proxyReq.host + proxyReq.path);
          },
          onError: function(err, req, res) {
            console.log('Proxy error:', err);
          }
        }
      ]
    },
  };

  return config;
};
