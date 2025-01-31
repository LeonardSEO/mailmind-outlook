const webpack = require('webpack');
const { merge } = require('webpack-merge');
const commonConfig = require('./webpack.common.js');

module.exports = merge(commonConfig, {
    mode: 'production',
    devtool: 'source-map',

    output: {
        publicPath: 'https://leonardseo.github.io/mailmind-outlook/'
    },

    performance: {
        hints: "warning"
    },

    optimization: {
        minimize: true,
        moduleIds: 'deterministic',
        runtimeChunk: 'single',
        splitChunks: {
            cacheGroups: {
                vendor: {
                    test: /[\\/]node_modules[\\/]/,
                    name: 'vendors',
                    chunks: 'all',
                },
            },
        },
    }
});

