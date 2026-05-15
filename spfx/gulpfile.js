'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

// Allow Chrome Private Network Access from SharePoint Online to localhost
build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    if (generatedConfiguration.devServer) {
      const originalBefore = generatedConfiguration.devServer.before;
      generatedConfiguration.devServer.before = (app, server, compiler) => {
        app.use((req, res, next) => {
          res.setHeader('Access-Control-Allow-Private-Network', 'true');
          next();
        });
        if (originalBefore) originalBefore(app, server, compiler);
      };
    }
    return generatedConfiguration;
  }
});

build.initialize(require('gulp'));
