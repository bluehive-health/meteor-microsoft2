Package.describe({
  name: 'bluehive-health:microsoft2',
  version: '1.0.2',
  summary: 'An implementation of the Microsoft OAuth flow using the v2.0 authorization endpoint.',
  git: 'https://github.com/bluehive-health/meteor-microsoft2',
  documentation: 'README.md'
});

Package.onUse(function(api) {
  api.versionsFrom('2.12');
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