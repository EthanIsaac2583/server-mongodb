const path = require('path');

let dependencies = Object.keys(require('./package.json').dependencies).reduce(function (acc, cur) {
    acc[cur] = cur
    return acc
}, new Object());

module.exports = {
    target: 'node',
    entry: {
        index: './src/index.js',
    },
    externalsPresets: { node: true },
    externals: dependencies,
    mode: 'production',
    output: {
        filename: '[name].js',
        path: path.resolve(__dirname, 'dist'),
        libraryTarget: 'commonjs2',
    },
    optimization: {
        minimize: false,
    },
    devServer: {
        // contentBase
        static : {
            directory : path.join(__dirname, "/")
        },
        headers: {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, PATCH, OPTIONS",
            "Access-Control-Allow-Headers": "X-Requested-With, content-type, Authorization"
        },
        port: 3005,
        devMiddleware: {
            publicPath: "https://localhost:3000/dist/",
        },
        hot: "only",
    },
    stats: { warnings: false },
};