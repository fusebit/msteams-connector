const Url = require('url');
const Sdk = require('../sdk'); // TODO @fusebit/add-on-sdk
const Mock = require('mock-http');
const Express = require('express');
const { createApp } = require('./app');
const { FusebitBot } = require('./FusebitBot');

// See https://github.com/fusebit/samples/blob/master/express/index.js#L6
Object.setPrototypeOf(Object.getPrototypeOf(Express.response), Mock.Response.prototype);
Object.setPrototypeOf(Object.getPrototypeOf(Express.request), Mock.Request.prototype);

exports.FusebitBot = FusebitBot;

exports.createMicrosoftTeamsConnector = (vendorBot) => {
    // Create Express app that exposes endpoints to receive notifications from Teams, handle vendor authorization, 
    // and sending of notifications. Teams notifications are handled by the vendor's bot. 
    const app = createApp(vendorBot);

    // Return a Fusebit handler that creates a mock HTTP request/response and hands the processing over to an Express app
    return (ctx, cb) => { 
        Sdk.debug('HTTP REQUEST', ctx.method, ctx.url, ctx.headers, ctx.body);

        ctx.storage = Sdk.getStorageClient(ctx);

        let url = ctx.url.split('/');
        url.splice(0,5);
        url = Url.parse('/' + url.join('/'));
        url.query = ctx.query;
        url = Url.format(url);

        const body = ctx.body ? Buffer.from(JSON.stringify(ctx.body)) : undefined;
        if (body) {
            ctx.headers['content-length'] = body.length;
        }

        let req = new Mock.Request({
            url,
            method: ctx.method, 
            headers: ctx.headers,
            buffer: body,
        }); 
        req.fusebit = ctx;

        let responseFinished;
        let res = new Mock.Response({
            onEnd: () => {
                if (responseFinished) return;
                responseFinished = true;
                timeout = undefined;
                const responseBody = (res._internal.buffer || Buffer.from('')).toString('utf8');
                Sdk.debug('HTTP RESPONSE', res.statusCode, responseBody);
                process.nextTick(() => cb(null, {  
                    body: responseBody,
                    bodyEncoding: 'utf8',
                    headers: res._internal.headers,
                    status: res.statusCode,
                }));
            }
        });

        app.handle(req, res);

        return;
    };
}
