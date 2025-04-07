const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const webpack = require('webpack');

module.exports = (env) => {
  const isLocalMode = env && env.local;
  
  const config = {
    entry: {
      index: isLocalMode ? ['./src/polyfills/process.js', './src/index.tsx'] : './src/index.tsx',
      commands: './src/commands.ts'
    },
    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: '[name].js',
      clean: true
    },
    resolve: {
      extensions: ['.ts', '.tsx', '.js', '.jsx'],
      alias: {
        // Use mock module for Salesforce in local mode
        '../services/salesforce': isLocalMode ? 
          path.resolve(__dirname, 'src/services/salesforce-mock.ts') : 
          path.resolve(__dirname, 'src/services/salesforce.ts')
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
          use: 'babel-loader',
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
          NODE_ENV: isLocalMode ? 'development' : 'production'
        })
      }),
      new webpack.ProvidePlugin({
        process: isLocalMode ? './src/polyfills/process.js' : 'process/browser'
      }),
      new HtmlWebpackPlugin({
        template: isLocalMode ? './src/local.html' : './src/index.html',
        filename: 'index.html',
        chunks: ['index']
      }),
      new HtmlWebpackPlugin({
        template: './src/commands.html',
        filename: 'commands.html',
        chunks: ['commands']
      })
    ],
    devServer: {
      static: {
        directory: path.join(__dirname, 'dist'),
      },
      headers: {
        'Access-Control-Allow-Origin': '*'
      },
      server: {
        type: 'https',
        options: {
          key: './certs/server.key',
          cert: './certs/server.crt',
        }
      },
      port: isLocalMode ? 3001 : 3000
    }
  };
  
  return config;
}; 