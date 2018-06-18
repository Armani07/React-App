'use strict';
var moment = require('moment-timezone');
const electron = require('electron');
const app = electron.app;

let winston = require('winston');
var logger = new (winston.Logger)({
    level: 'info',
    exitOnError: false,
    transports: [
        new (winston.transports.Console)({
            humanReadableUnhandledException: true,
            prettyPrint: true,
            handleExceptions: true,
            timestamp: function () {
                //return Date.now();
                return moment().format('YYYY-MM-DD hh:mm:ss:ms')
            }
        }),
        new (winston.transports.File)({
            filename: process.env.APPDATA + '/Visual Database Assurer/somefile.log',
            humanReadableUnhandledException: true,
            handleExceptions: true,
            prettyPrint: true,
            timestamp: function () {
                //return Date.now();
                return moment().format('YYYY-MM-DD hh:mm:ss:ms')
            }
        })
    ]
});
module.exports = logger;