'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
const merge = require('webpack-merge');
const webpack = require('webpack');
const argv = require('yargs').argv;
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

build.configureWebpack.setConfig({
  additionalConfiguration: function (config) {
    var isProduction = (argv.master === undefined) ? false : true;
    let defineOptions;

    // Production
    if (isProduction) {
      console.log('***********    Applying production environment settings to webpack *********************');
      defineOptions = {
        "Environment": JSON.stringify("stortinget"),
        "HttpCreateTeamUrl": JSON.stringify("http://localhost:7071/api/getTeamConfig")
      }
    }
    // Test/development
    else {
      console.log('***********    Applying development environment settings to webpack *********************');
      defineOptions = {
        "Environment": JSON.stringify("utvstortinget"),
        "HttpCreateTeamUrl": JSON.stringify("http://localhost:7071/api/getTeamConfig")
      }
    }

    return merge(config, {
      plugins: [
        new webpack.DefinePlugin(defineOptions)
      ]
    });
  }
});

build.initialize(gulp);