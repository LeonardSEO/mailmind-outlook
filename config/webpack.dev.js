const fs = require('fs');
const path = require('path');
const webpack = require('webpack');
const { merge } = require('webpack-merge');
const commonConfig = require('./webpack.common.js');

const devCertsPath = path.resolve(process.env.USERPROFILE, '.office-addin-dev-certs');

module.exports = merge(commonConfig, {
    mode: 'development',
    devtool: 'eval-source-map',
    devServer: {
        static: {
            directory: path.resolve('dist'),
        },
        hot: true,
        server: {
            type: 'https',
            options: {
                key: fs.readFileSync(path.resolve(devCertsPath, 'localhost.key')),
                cert: fs.readFileSync(path.resolve(devCertsPath, 'localhost.crt'))
            }
        },
        compress: true,
        client: {
            overlay: {
                warnings: false,
                errors: true
            }
        },
        port: 3000,
        historyApiFallback: true,
        headers: {
            "Access-Control-Allow-Origin": "*"
        }
    }
});
