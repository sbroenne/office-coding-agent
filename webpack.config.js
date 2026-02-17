/* eslint-disable @typescript-eslint/no-require-imports */
const path = require('path');
const webpack = require('webpack');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const MiniCssExtractPlugin = require('mini-css-extract-plugin');

const devCerts = require('office-addin-dev-certs');

module.exports = async (env, options) => {
  const isDev = options.mode === 'development';

  const devServerOptions = isDev
    ? {
        devServer: {
          hot: true,
          port: 3000,
          server: {
            type: 'https',
            options: await devCerts.getHttpsServerOptions(),
          },
          headers: {
            'Access-Control-Allow-Origin': '*',
          },
          allowedHosts: 'all',
          client: {
            overlay: false, // Disable error overlay — Office Add-in WebView generates cross-origin "Script error." noise
          },
        },
      }
    : {};

  return {
    ...devServerOptions,
    entry: {
      taskpane: './src/taskpane/index.tsx',
    },
    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: '[name].[contenthash].js',
      clean: true,
    },
    resolve: {
      extensions: ['.ts', '.tsx', '.js', '.jsx'],
      alias: {
        '@': path.resolve(__dirname, 'src'),
      },
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: {
            loader: 'ts-loader',
            options: {
              // Only type-check files that webpack actually bundles —
              // prevents ts-loader from compiling test files listed in tsconfig.include
              onlyCompileBundledFiles: true,
              compilerOptions: {
                noEmit: false,
                types: ['office-js', 'office-runtime', 'node'],
              },
            },
          },
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: [
            isDev ? 'style-loader' : MiniCssExtractPlugin.loader,
            'css-loader',
            'postcss-loader',
          ],
        },
        {
          test: /\.md$/,
          type: 'asset/source',
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: './src/taskpane/taskpane.html',
        filename: 'taskpane.html',
        chunks: ['taskpane'],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: 'assets',
            to: 'assets',
            noErrorOnMissing: true,
          },
        ],
      }),
      new webpack.DefinePlugin({
        'process.env.AZURE_OPENAI_ENDPOINT': JSON.stringify(
          process.env.AZURE_OPENAI_ENDPOINT || ''
        ),
        'process.env.AZURE_OPENAI_API_KEY': JSON.stringify(process.env.AZURE_OPENAI_API_KEY || ''),
      }),
      ...(isDev ? [] : [new MiniCssExtractPlugin({ filename: '[name].[contenthash].css' })]),
    ],
    // Office task panes ship a single bundled entry loaded in WebView; webpack's generic web perf
    // asset-size hints are noisy here and not actionable for our sideloaded add-in packaging model.
    performance: {
      hints: false,
    },
    devtool: isDev ? 'source-map' : false,
  };
};
