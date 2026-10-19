'use strict';
const serverless = require('serverless-http');
const app = require('../../server.js');

const serverlessHandler = serverless(app);

module.exports.handler = async (event, context) => {
  // Normalize request path so Express route matching always sees /api/...
  if (event.path && event.path.startsWith('/.netlify/functions/api')) {
    const sub = event.path.replace(/^\/\.netlify\/functions\/api/, '');
    if (!sub.startsWith('/api')) {
      event.path = '/api' + (sub || '');
    } else {
      event.path = sub;
    }
  }
  return await serverlessHandler(event, context);
};
