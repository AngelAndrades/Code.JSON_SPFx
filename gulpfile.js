'use strict';

const path = require('path');
const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// clean up React error messages
build.configureWebpack.mergeConfig({
    additionalConfiguration: (generatedConfiguration) => {
        if (!generatedConfiguration.resolve.alias) {
            generatedConfiguration.resolve.alias = {};
        }
        generatedConfiguration.resolve.alias['@utils'] = path.resolve(__dirname, 'lib/utils');
        generatedConfiguration.resolve.alias['@src'] = path.resolve(__dirname, 'lib');

        generatedConfiguration.externals.splice(generatedConfiguration.externals.indexOf('react'), 1);
        generatedConfiguration.externals.splice(generatedConfiguration.externals.indexOf('react-dom'), 1);

        return generatedConfiguration;
    }
});

build.initialize(require('gulp'));

/*
 *  Toolchain customization to keep build numbers synced using:
 *  npm version (major|minor|patch)
 * 
 *  syncs the package.json with the /config/package-solution.json
*/
gulp.task('version-sync', _ => {
    // import gulp utilities to write error messages
    const gutil = require('gulp-util');

    // import file system utilities form nodeJS
    const fs = require('fs');

    // read package.json
    var pkgConfig = require('./package.json');

    // read configuration of web part solution file
    var pkgSolution = require('./config/package-solution.json');

    // log old version
    gutil.log('Old Version:\t' + pkgSolution.solution.version);

    // Generate new MS compliant version number
    var newVersionNumber = pkgConfig.version.split('-')[0] + '.0';

    // assign newly generated version number to web part version
    pkgSolution.solution.version = newVersionNumber;

    // log new version
    gutil.log('New Version:\t' + pkgSolution.solution.version);

    // write changed package-solution file
    fs.writeFile('./config/package-solution.json', JSON.stringify(pkgSolution, null, 2), (err) => {
        gutil.log('Error updating file: ' + err);
    });
});
