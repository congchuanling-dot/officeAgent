const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = (_env, argv) => {
  const isDev = argv.mode === 'development';

  return {
    entry: './src/taskpane.tsx',
    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: 'taskpane.js',
      clean: true,
    },
    resolve: {
      extensions: ['.tsx', '.ts', '.js', '.jsx'],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: 'ts-loader',
          exclude: /node_modules/,
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: './src/taskpane.html',
        filename: 'taskpane.html',
        inject: 'body',
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: 'manifest.xml', to: 'manifest.xml' },
        ],
      }),
    ],
    devServer: {
      static: './dist',
      https: true,
      port: 3000,
      hot: false,
    },
    devtool: isDev ? 'eval-source-map' : false,
  };
};
