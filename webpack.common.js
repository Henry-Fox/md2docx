const path = require('path');

module.exports = {
  entry: {
    app: './js/app.js',
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    clean: true,
    filename: './js/app.js',
  },
  module: {
    rules: [
      {
        test: /\.json$/,
        type: 'json',
      }
    ]
  },
  resolve: {
    extensions: ['.js', '.json']
  }
};
