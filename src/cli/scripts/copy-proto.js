'use strict';

const fs = require('node:fs');
const path = require('node:path');
const source = path.resolve(__dirname, '../../proto/netoffice.proto');
const destination = path.resolve(__dirname, '../proto/netoffice.proto');
fs.mkdirSync(path.dirname(destination), { recursive: true });
fs.copyFileSync(source, destination);
