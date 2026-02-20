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
 * Acquire an Entra ID bearer token at build time when no API key is set.
 * The token is baked into the bundle so E2E tests (running in Excel's WebView,
 * where @azure/identity is unavailable) can authenticate with Azure AI.
 */
async function acquireBearerToken() {
  if (process.env.FOUNDRY_API_KEY || !process.env.FOUNDRY_ENDPOINT) return '';
  try {
    const { DefaultAzureCredential } = require('@azure/identity');
    const credential = new DefaultAzureCredential();
    const response = await credential.getToken('https://cognitiveservices.azure.com/.default');
    if (response?.token) {
      console.log('✓ Acquired Entra ID bearer token for E2E AI tests');
      return response.token;
    }
  } catch (err) {
    console.warn('⚠ Could not acquire Entra ID token:', err.message);
  }
  return '';
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const bearerToken = await acquireBearerToken();
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
        'process.env.FOUNDRY_ENDPOINT': JSON.stringify(process.env.FOUNDRY_ENDPOINT || ''),
        'process.env.FOUNDRY_API_KEY': JSON.stringify(process.env.FOUNDRY_API_KEY || ''),
        'process.env.FOUNDRY_MODEL': JSON.stringify(process.env.FOUNDRY_MODEL || 'gpt-5.2-chat'),
        'process.env.FOUNDRY_BEARER_TOKEN': JSON.stringify(bearerToken),
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
