const path = require('path');
const GasPlugin = require('gas-webpack-plugin');
const TerserPlugin = require('terser-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const InlineChunkHtmlPlugin = require('react-dev-utils/InlineChunkHtmlPlugin');

const config = {
  entry: {
    appscript: './src/index.ts',
    client: './src/client/client.ts',
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    clean: true,
    filename: '[name].js',
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
    // Generates an `Page.html` file with the <script> injected.
    new HtmlWebpackPlugin({
      inject: true,
      template: path.resolve('public/Page.html'),
      filename: 'Page.html',
      inlineSource: '.(js|css)$',
      cache: false,
      chunks: ['client'],
    }),
    // Inlines chunks with `runtime` in the name
    new InlineChunkHtmlPlugin(HtmlWebpackPlugin, [/client/]),
    new GasPlugin({ include: ['src/**/*.ts'] }),
  ],
};

module.exports = config;
