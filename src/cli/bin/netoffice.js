#!/usr/bin/env node
'use strict';

const { parseArguments, usage } = require('../lib/arguments');
const fs = require('node:fs');
const path = require('node:path');
const { execFile } = require('node:child_process');

const endpoint = '127.0.0.1:50051';

// Fixed script: no command-line title or other user text is evaluated by PowerShell.
const launchScript = `
$ErrorActionPreference = 'Stop'
$paths = @(
 'Registry::HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\POWERPNT.EXE',
 'Registry::HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\POWERPNT.EXE',
 'Registry::HKEY_LOCAL_MACHINE\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\App Paths\\POWERPNT.EXE'
)
foreach ($key in $paths) {
 if (Test-Path $key) {
  $exe = (Get-Item $key).GetValue('')
  if ($exe -and (Test-Path -LiteralPath $exe)) {
   Start-Process -FilePath $exe
   exit 0
  }
 }
}
try {
 $app = New-Object -ComObject PowerPoint.Application
 $app.Visible = -1
 [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)
} catch {
 throw 'PowerPoint is not installed or its PowerPoint.Application COM registration is unavailable. Install desktop Microsoft PowerPoint and register the NetOffice native add-in.'
}
`;

function launchPowerPoint(deadline) {
  return new Promise((resolve, reject) => {
    const executable = path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
    execFile(executable, ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', launchScript], {
      windowsHide: true,
      timeout: Math.max(1, deadline - Date.now()),
      maxBuffer: 1024 * 1024
    }, (error, stdout, stderr) => {
      if (error) reject(new Error(`Unable to launch PowerPoint: ${stderr.trim() || error.message}`));
      else resolve();
    });
  });
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  if (options.help) { process.stdout.write(usage); return; }
  if (options.command === 'powerpoint launch' && process.platform !== 'win32') {
    throw new Error('powerpoint launch is supported only on Windows with desktop Microsoft PowerPoint installed.');
  }
  const grpc = require('@grpc/grpc-js');
  const protoLoader = require('@grpc/proto-loader');
  const packagedProto = path.resolve(__dirname, '../proto/netoffice.proto');
  const sharedProto = path.resolve(__dirname, '../../proto/netoffice.proto');
  const protoPath = fs.existsSync(sharedProto) ? sharedProto : packagedProto;
  const definition = protoLoader.loadSync(protoPath, { keepCase: true, defaults: true, oneofs: true });
  const { netoffice } = grpc.loadPackageDefinition(definition);
  const client = new netoffice.PowerPoint(endpoint, grpc.credentials.createInsecure(), { 'grpc.enable_retries': 0 });
  const deadline = Date.now() + options.timeout;
  const rpc = (method, request, end = deadline) => new Promise((resolve, reject) => {
    client[method](request, { deadline: new Date(end) }, (error, reply) => error ? reject(error) : resolve(reply));
  });
  const ready = () => new Promise((resolve, reject) => {
    client.waitForReady(new Date(deadline), error => error ? reject(error) : resolve());
  });
  try {
    if (options.command === 'powerpoint launch') {
      let status;
      try {
        status = await rpc('getStatus', {}, Math.min(deadline, Date.now() + 500));
      } catch (error) {
        if (![grpc.status.UNAVAILABLE, grpc.status.DEADLINE_EXCEEDED].includes(error.code)) throw error;
        if (Date.now() >= deadline) throw error;
        await launchPowerPoint(deadline);
        await ready();
        status = await rpc('getStatus', {});
      }
      console.log(`PowerPoint ready (PID ${status.process_id}).`);
    } else {
      await ready();
      if (options.command === 'presentation new') {
        const reply = await rpc('newPresentation', { title: options.title });
        console.log(`Created presentation ${JSON.stringify(reply.name)} (${reply.slide_count} slide${reply.slide_count === 1 ? '' : 's'}).`);
      } else {
        await rpc('setSlideTitle', { slide_index: options.slide, text: options.title });
        console.log(`Updated slide ${options.slide} title to ${JSON.stringify(options.title)}.`);
      }
    }
  } catch (error) {
    const status = Object.keys(grpc.status).find(name => grpc.status[name] === error.code);
    const unavailable = error.code === grpc.status.UNAVAILABLE || error.code === grpc.status.DEADLINE_EXCEEDED || (!Number.isInteger(error.code) && Date.now() >= deadline);
    const detail = error.details || error.message;
    throw new Error(`${status ? `${status}: ` : ''}${detail}${unavailable ? '\nEnsure PowerPoint is running and the NetOffice native add-in is registered, enabled, and listening at ' + endpoint + '. Use netoffice powerpoint launch on Windows or increase --timeout.' : ''}`);
  } finally {
    client.close();
  }
}

main().catch(error => {
  console.error(`netoffice: ${error.message}`);
  process.exitCode = 1;
});
