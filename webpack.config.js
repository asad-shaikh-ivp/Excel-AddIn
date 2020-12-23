const devCerts = require("office-addin-dev-certs");
const CleanWebpackPlugin = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const CustomFunctionsMetadataPlugin = require("custom-functions-metadata-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const fs = require("fs");
const webpack = require("webpack");


const API_URL = {
  uat: 'https://cfrms-uat.azurewebsites.net',
  production: 'https://cfrms.azurewebsites.net',
  development: 'https://localhost:5001'
}
const ADDIN_URL = {
  uat: 'https://cfrms-web-uat.azurewebsites.net/exceladdin',
  production: 'https://cfrms-web.azurewebsites.net/exceladdin',
  development: 'https://localhost:3000',
}
const AD_AUTH_CLIENT_ID = {
  uat: '981b9327-5c3b-4cbc-ae59-f52442fff09c',
  production: '981b9327-5c3b-4cbc-ae59-f52442fff09c',
  development: '981b9327-5c3b-4cbc-ae59-f52442fff09c'
}
const AD_AUTH_TENANT_ID = {
  uat: '62c5eb46-d129-44dd-91fd-47d9e6a17d69',
  production: '62c5eb46-d129-44dd-91fd-47d9e6a17d69',
  development: '62c5eb46-d129-44dd-91fd-47d9e6a17d69'
}


module.exports = async (env, options) => {
  const dev = options.mode === "development";
  var environment = options.mode  === 'production' ? 'production' : 'development';

console.log(env);

  let deploymentUrlPaths = 'development' ;
  if(env && env.hasOwnProperty("deploymentUrlPaths") && env.deploymentUrlPaths){
    deploymentUrlPaths = env.deploymentUrlPaths;
  }
  console.log(deploymentUrlPaths);
  console.log(AD_AUTH_TENANT_ID[deploymentUrlPaths]);
  console.log(AD_AUTH_CLIENT_ID[deploymentUrlPaths]);
  console.log(ADDIN_URL[deploymentUrlPaths]);
  console.log(API_URL[deploymentUrlPaths]);

  //var environment = 'development';
  const config = {
    devtool: "source-map",
    entry: {
      vendor: [
        'react',
        'react-dom',
        'core-js',
        'office-ui-fabric-react'
      ],
      app: './src/index.tsx',
      functions: "./src/functions/functions.ts",
      polyfill: "@babel/polyfill",
      taskpane: "./src/taskpane/taskpane.ts",
      commands: "./src/commands/commands.ts",
      login: "./login/login.ts",
      logout: "./logout/logout.ts"
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"]
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: "babel-loader"
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: "ts-loader"
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader"
        },
        {
          test: /\.(png|jpg|jpeg|gif)$/,
          use: "file-loader"
        },
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader']
        },
        {
          test: /\.less$/,
          use: ['style-loader', 'css-loader', 'less-loader']
        },
        {
          test: /\.(png|jpe?g|gif|svg|woff|woff2|ttf|eot|ico)$/,
          use: {
            loader: 'file-loader',
            query: {
              name: 'assets/[name].[ext]'
            }
          }
        }
      ]
    },
    plugins: [
      new CleanWebpackPlugin({
        cleanOnceBeforeBuildPatterns: dev ? [] : ["**/*"]
      }),

      new webpack.DefinePlugin({
        'API_URL': JSON.stringify(API_URL[deploymentUrlPaths]),
        'ADDIN_URL': JSON.stringify(ADDIN_URL[deploymentUrlPaths]),
        'AD_AUTH_CLIENT_ID': JSON.stringify(AD_AUTH_CLIENT_ID[deploymentUrlPaths]),
        'AD_AUTH_TENANT_ID': JSON.stringify(AD_AUTH_TENANT_ID[deploymentUrlPaths])
      }),

      new HtmlWebpackPlugin({
        //title: 'Capital Four Research Management System',
        filename: 'index.html',
        template: './src/index.html',
        chunks: ['app', 'vendor', 'polyfill']
      }),
      new CustomFunctionsMetadataPlugin({
        output: "functions.json",
        input: "./src/functions/functions.ts"
      }),
      new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"]
      }),
      new CopyWebpackPlugin([
        {
          to: "taskpane.css",
          from: "./src/taskpane/taskpane.css"
        }
      ]),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      }), new HtmlWebpackPlugin({
        //title: 'Capital Four Research Management System',
        filename: 'login/login.html',
        template: './login/login.html',
        chunks: ['login']
      }),
      new HtmlWebpackPlugin({
        //title: 'Capital Four Research Management System',
        filename: 'logout/logout.html',
        template: './logout/logout.html',
        chunks: ['logout']
      }),
      new HtmlWebpackPlugin({
        //title: 'Capital Four Research Management System',
        filename: 'logoutcomplete/logoutcomplete.html',
        template: './logoutcomplete/logoutcomplete.html',
        chunks: ['logoutcomplete']
      }),
      new HtmlWebpackPlugin({
        title: 'Capital Four Research Management System',
        filename: "Dialog/dialog.html",
        template: "./Dialog/dialog.html",
        chunks: ["dialog"]
      }),
      new HtmlWebpackPlugin({
        title: 'Capital Four Research Management System',
        filename: "Dialog/deleteAI.html",
        template: "./Dialog/deleteAI.html",
        chunks: ["delete"]
      }),
      new HtmlWebpackPlugin({
        title: 'Capital Four Research Management System',
        filename: "Dialog/success.html",
        template: "./Dialog/success.html",
        chunks: ["success"]
      }),
      new HtmlWebpackPlugin({
        title: 'Capital Four Research Management System',
        filename: "ErrorDialog/dialog.html",
        template: "./ErrorDialog/dialog.html",
        chunks: ["errordialog"]
      }),
      new HtmlWebpackPlugin({
        title: 'Capital Four Research Management System',
        filename: "ErrorDialog/fetchError.html",
        template: "./ErrorDialog/fetchError.html",
        chunks: ["fetchError"]
      }),
      new HtmlWebpackPlugin({
        title: 'Capital Four Research Management System',
        filename: "ErrorDialog/excelRunError.html",
        template: "./ErrorDialog/excelRunError.html",
        chunks: ["excelRunError"]
      }),
      new HtmlWebpackPlugin({
        title: 'Capital Four Research Management System',
        filename: "ErrorDialog/getFileError.html",
        template: "./ErrorDialog/getFileError.html",
        chunks: ["getFileError"]
      }),
      new HtmlWebpackPlugin({
        template: './ErrorDialog/stringError.html',
        filename: 'ErrorDialog/dialog.html',
        chunks: ['stringError']
      }),
      new HtmlWebpackPlugin({
        template: './ErrorDialog/invalidMnemonicsInModel.html',
        filename: 'ErrorDialog/invalidMnemonicsInModel.html',
        chunks: ['invalidMnemonicsInModel']
      }),
      new CopyWebpackPlugin([
        {
          to: '.',
          from: "./ErrorDialog/"
        }
      ]),
      new HtmlWebpackPlugin({
        template: './ErrorDialog/dirtySheet.html',
        filename: 'ErrorDialog/dirtySheet.html',
        chunks: ['dirtySheet']
      }),
      new HtmlWebpackPlugin({
        template: './ErrorDialog/invalidIdentifier.html',
        filename: 'ErrorDialog/invalidIdentifier.html',
        chunks: ['invalidIdentifier']
      }),
      new CopyWebpackPlugin([
        {
          from: './assets',
          ignore: ['*.scss'],
          to: 'assets',
        }
      ]),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"]
      })
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*"
      },
      https: environment === 'development' ? await devCerts.getHttpsServerOptions() : null,
      port: 3000,
      host: "0.0.0.0",
      compress: true,
      disableHostCheck: true
    }
  };

  return config;
};
