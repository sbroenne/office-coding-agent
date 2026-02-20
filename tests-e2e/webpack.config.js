/* eslint-disable @typescript-eslint/no-require-imports */

/**
 * Webpack configuration for E2E Tests
 *
 * Builds a standalone test taskpane that runs Excel command tests
 * inside a real Excel instance and reports results to a test server.
 * Served on port 3001 to avoid conflicting with the dev add-in on 3000.
 */

const devCerts = require('office-addin-dev-certs');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const webpack = require('webpack');
const path = require('path');

// Load .env from project root for AI credentials
try {
  require('dotenv').config({ path: path.resolve(__dirname, '../.env') });
} catch {
  /* dotenv optional */
}

/**
 * Acquire an Entra ID bearer token at build time (legacy - kept for compatibility).
 */
async function acquireBearerToken() {
  return '';
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const config = {
    devtool: 'source-map',
    entry: {
      taskpane: path.resolve(__dirname, 'src/test-taskpane.ts'),
      functions: path.resolve(__dirname, 'src/functions.ts'),
    },
    output: {
      path: path.resolve(__dirname, 'dist'),
      clean: true,
    },
    resolve: {
      extensions: ['.ts', '.tsx', '.html', '.js'],
      alias: {
        '@': path.resolve(__dirname, '../src'),
      },
      fallback: {
        child_process: false,
        fs: false,
        process: false,
      },
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: {
            loader: 'ts-loader',
            options: {
              transpileOnly: true,
              configFile: path.resolve(__dirname, '../tsconfig.json'),
              compilerOptions: {
                noEmit: false,
              },
            },
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: 'html-loader',
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: 'asset/resource',
          generator: {
            filename: 'assets/[name][ext][query]',
          },
        },
        {
          test: /\.md$/,
          type: 'asset/source',
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: 'test-taskpane.html',
        template: path.resolve(__dirname, 'src/test-taskpane.html'),
        chunks: ['taskpane'],
        inject: 'body',
        scriptLoading: 'blocking',
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: path.resolve(__dirname, '../assets'),
            to: 'assets',
            noErrorOnMissing: true,
          },
          {
            from: path.resolve(__dirname, 'src/functions.json'),
            to: 'functions.json',
          },
        ],
      }),
      new webpack.DefinePlugin({
        'process.env.COPILOT_SERVER_URL': JSON.stringify(
          process.env.COPILOT_SERVER_URL || 'wss://localhost:3000/api/copilot'
        ),
      }),
    ],
    devServer: {
      static: {
        directory: path.resolve(__dirname, 'dist'),
        publicPath: '/',
      },
      headers: {
        'Access-Control-Allow-Origin': '*',
      },
      server: {
        type: 'https',
        options:
          env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },
      port: 3001,
      client: {
        overlay: false,
      },
    },
  };

  return config;
};
