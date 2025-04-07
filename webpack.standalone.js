const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const webpack = require('webpack');

module.exports = {
  mode: 'development',
  entry: ['./src/polyfills/process.js', './src/standalone.tsx'],
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: 'bundle.js',
    clean: true
  },
  resolve: {
    extensions: ['.ts', '.tsx', '.js', '.jsx'],
    alias: {
      // Use mock implementations
      '../services/salesforce': path.resolve(__dirname, 'src/services/salesforce-mock.ts'),
      '../services/office': path.resolve(__dirname, 'src/services/office-mock.ts')
    },
    fallback: {
      // Provide empty mocks for Node.js core modules
      "fs": false,
      "path": false,
      "os": false,
      "assert": false,
      "stream": false,
      "timers": false,
      "util": false,
      "buffer": false,
      "crypto": false,
      "querystring": false,
      "http": false,
      "https": false,
      "url": false,
      "zlib": false,
      "net": false,
      "tls": false
    }
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: [
          {
            loader: 'babel-loader',
            options: {
              presets: [
                '@babel/preset-env',
                '@babel/preset-react',
                '@babel/preset-typescript'
              ]
            }
          },
          {
            loader: 'ts-loader',
            options: {
              configFile: 'tsconfig.simple.json',
              transpileOnly: true
            }
          }
        ],
        exclude: /node_modules/
      },
      {
        test: /\.css$/,
        use: ['style-loader', 'css-loader']
      }
    ]
  },
  plugins: [
    new webpack.DefinePlugin({
      'process.env': JSON.stringify({
        NODE_ENV: 'development',
        SF_LOGIN_URL: 'https://login.salesforce.com',
        SF_CLIENT_ID: 'mock_client_id',
        SF_CLIENT_SECRET: 'mock_client_secret',
        SF_REDIRECT_URI: 'https://localhost:8000/oauth/callback'
      })
    }),
    new HtmlWebpackPlugin({
      template: './src/standalone.html',
      filename: 'index.html'
    }),
    new HtmlWebpackPlugin({
      template: './public/oauth-callback.html',
      filename: 'oauth-callback.html',
      inject: false
    })
  ],
  devtool: 'inline-source-map',
  devServer: {
    static: {
      directory: path.join(__dirname, 'dist'),
    },
    port: 8000,
    open: true,
    hot: true,
    historyApiFallback: {
      rewrites: [
        { from: /^\/oauth-callback$/, to: '/oauth-callback.html' }
      ]
    }
  }
}; 