Package.describe({
  name: 'ghobbs:microsoft2',
  version: '1.0.2',
  summary: 'An implementation of the Microsoft OAuth flow using the v2.0 authorization endpoint.',
  git: 'https://github.com/gwhobbs/meteor-microsoft',
  documentation: 'README.md'
});

Package.onUse(function(api) {
  api.versionsFrom('1.2.1');
  api.use('ecmascript');
  api.use('templating', 'client');
  api.use('random', 'client');
  api.use('oauth2', ['client', 'server']);
  api.use('oauth', ['client', 'server']);
  api.use('http', 'server');
  api.use('service-configuration', ['client', 'server']);

  api.export('Microsoft');

  api.addFiles('client/configure.html', 'client');
  api.addFiles([
    'server/server.js',
    'server/http.js',
    'server/tokens.js',
    'server/graph/graph.js',
    'server/graph/classes.js',
  ], 'server');
  api.addFiles(['client/client.js', 'client/configure.js'], 'client');
});

Package.onTest(function(api) {
  api.use('tinytest');
  api.use('ecmascript');
  api.use('ghobbs:microsoft2');

  // Tests will follow soon!
  api.addFiles([]);
});
