const { merge } = require('webpack-merge');
const common = require('./webpack.common.js');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyPlugin = require('copy-webpack-plugin');

module.exports = merge(common, {
  mode: 'production',
  plugins: [
    new HtmlWebpackPlugin({
      template: './index.html',
    }),
    new CopyPlugin({
      patterns: [
        { from: 'css', to: 'css' },
        { from: 'icon.svg', to: 'icon.svg' },
        { from: 'icon.png', to: 'icon.png' },
        { from: 'favicon.ico', to: 'favicon.ico' },
        { from: 'site.webmanifest', to: 'site.webmanifest' },
        { from: 'robots.txt', to: 'robots.txt' },
        { from: 'sitemap.xml', to: 'sitemap.xml' },
        { from: 'llms.txt', to: 'llms.txt' },
        { from: 'img', to: 'img' },
      ],
    }),
  ],
});
