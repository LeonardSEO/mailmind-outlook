const webpack = require('webpack');
const path = require('path');
const package = require('../package.json');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const MiniCssExtractPlugin = require('mini-css-extract-plugin');

const build = (() => {
    const timestamp = new Date().getTime();
    return {
        name: package.name,
        version: package.version,
        timestamp: timestamp,
        author: package.author
    };
})();

const entry = {
    app: './index.tsx',
    'function-file': '../function-file/function-file.ts'
};

const rules = [
    {
        test: /\.tsx?$/,
        use: 'ts-loader',
        exclude: /node_modules/
    },
    {
        test: /\.css$/,
        use: [
            MiniCssExtractPlugin.loader,
            'css-loader',
            {
                loader: 'postcss-loader',
                options: {
                    postcssOptions: {
                        plugins: [
                            ['autoprefixer']
                        ]
                    }
                }
            }
        ]
    },
    {
        test: /\.less$/,
        use: [
            MiniCssExtractPlugin.loader,
            'css-loader',
            {
                loader: 'postcss-loader',
                options: {
                    postcssOptions: {
                        plugins: [
                            ['autoprefixer']
                        ]
                    }
                }
            },
            'less-loader'
        ]
    },
    {
        test: /\.(png|jpe?g|gif|svg|woff|woff2|ttf|eot|ico)$/,
        type: 'asset/resource',
        generator: {
            filename: 'assets/[name][ext]'
        }
    }
];

const output = {
    path: path.resolve('dist'),
    publicPath: '/',
    filename: '[name].[contenthash].js',
    chunkFilename: '[id].[contenthash].chunk.js',
    clean: true
};

const WEBPACK_PLUGINS = [
    new webpack.NoEmitOnErrorsPlugin(),
    new webpack.BannerPlugin({ banner: `${build.name} v.${build.version} (${build.timestamp}) © ${build.author}` }),
    new webpack.DefinePlugin({
        ENVIRONMENT: JSON.stringify({
            build: build
        })
    })
];

module.exports = {
    context: path.resolve('./src'),
    entry,
    output,
    resolve: {
        extensions: ['.js', '.jsx', '.ts', '.tsx', '.scss', '.css', '.html']
    },
    module: {
        rules,
    },
    optimization: {
        splitChunks: {
            chunks: 'all',
            cacheGroups: {
                defaultVendors: {
                    test: /[\\/]node_modules[\\/]/,
                    name: 'vendors',
                    chunks: 'all',
                    priority: -10
                },
                default: {
                    minChunks: 2,
                    priority: -20,
                    reuseExistingChunk: true
                }
            }
        }
    },
    plugins: [
        ...WEBPACK_PLUGINS,
        new MiniCssExtractPlugin({
            filename: '[name].[contenthash].css',
            chunkFilename: '[id].[contenthash].css'
        }),
        new HtmlWebpackPlugin({
            title: 'outlook-addin-using-react-demo',
            filename: 'index.html',
            template: './index.html',
            chunks: ['app']
        }),
        new HtmlWebpackPlugin({
            title: 'outlook-addin-using-react-demo',
            filename: 'function-file/function-file.html',
            template: '../function-file/function-file.html',
            chunks: ['function-file']
        }),
        new CopyWebpackPlugin({
            patterns: [
                {
                    from: '../assets',
                    globOptions: {
                        ignore: ['*.scss']
                    },
                    to: 'assets'
                }
            ]
        })
    ]
};