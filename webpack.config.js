const path = require('path');
const GasPlugin = require('gas-webpack-plugin');
const TerserPlugin = require('terser-webpack-plugin');
const CopyPlugin = require('copy-webpack-plugin');

const config = {
  entry: './src/index.ts',

  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: 'bundle.js',
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        use: ['ts-loader'],
        exclude: /node_modules/,
      },
    ],
  },
  resolve: {
    extensions: ['.ts'],
  },
  optimization: {
    minimize: false,
    minimizer: [
      new TerserPlugin({
        terserOptions: {
          mangle: false,
          output: {
            comments: /@customFunction/i,
          },
        },
      }),
    ],
  },
  plugins: [
    new GasPlugin(),
    new CopyPlugin({
      patterns: [
        {
          from: 'public/*.html',
          to({ context, absoluteFileName }) {
            return '[name][ext]';
          },
        },
      ],
    }),
  ],
};

module.exports = config;
