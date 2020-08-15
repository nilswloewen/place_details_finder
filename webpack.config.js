const {CleanWebpackPlugin} = require('clean-webpack-plugin');
const CopyPlugin = require('copy-webpack-plugin');
const ExtractTextPlugin = require('extract-text-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const devCerts = require('office-addin-dev-certs');
const webpack = require('webpack');

module.exports = async (env, options) => {
  const dev = options.mode === 'development';
  const config = {
    devtool: 'source-map',
    entry: {
      vendor: [
        'react',
        'react-dom',
        'core-js',
        'office-ui-fabric-react',
      ],
      polyfill: 'babel-polyfill',
      taskpane: [
        'react-hot-loader/patch',
        './src/taskpane/index.js',
      ],
      commands: './src/commands/commands.js',
    },
    resolve: {
      extensions: ['.ts', '.tsx', '.html', '.js'],
    },
    module: {
      rules: [
        {
          test: /\.jsx?$/,
          use: [
            'react-hot-loader/webpack',
            'babel-loader',
          ],
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader'],
        },
        {
          test: /\.(png|jpe?g|gif|svg|woff|woff2|ttf|eot|ico)$/,
          use: {
            loader: 'file-loader',
            query: {
              name: 'assets/[name].[ext]',
            },
          },
        },
      ],
    },
    plugins: [
      new CleanWebpackPlugin(),
      new CopyPlugin({
        patterns: [
          { from: './src/taskpane/taskpane.css', to: 'taskpane.css' },
          { from: './assets', to: 'assets' },
        ],
      }),
      new ExtractTextPlugin('[name].[hash].css'),
      new HtmlWebpackPlugin({
        filename: 'taskpane.html',
        template: './src/taskpane/taskpane.html',
        chunks: ['taskpane', 'vendor', 'polyfill'],
      }),
      new HtmlWebpackPlugin({
        filename: 'commands.html',
        template: './src/commands/commands.html',
        chunks: ['commands'],
      }),
      new HtmlWebpackPlugin({
        filename: 'index.html',
        template: './src/taskpane/index.html',
      }),
      new webpack.ProvidePlugin({
        Promise: ['es6-promise', 'Promise'],
      }),
    ],
    devServer: {
      headers: {
        'Access-Control-Allow-Origin': '*',
      },
      https: (options.https !== undefined) ?
          options.https :
          await devCerts.getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
