#!/usr/bin/env node
'use strict';

const { parseArguments, usage } = require('../lib/arguments');
const path = require('node:path');
const { execFile } = require('node:child_process');

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
  const { Connection, deadlineError, retryable } = require('../lib/connection');
  const endpoint = `ws://127.0.0.1:${options.port}/devtools/application`;
  const deadline = Date.now() + options.timeout;
  let client;
  const connect = async end => {
    client = await Connection.connect(endpoint, end);
  };
  const status = async end => {
    const reply = await client.request('PowerPoint.getStatus', {}, end);
    if (!Number.isSafeInteger(reply.processId) || reply.processId < 1) throw new Error('Invalid server response: expected a positive processId.');
    return reply;
  };
  const ready = async requireStatus => {
    while (Date.now() < deadline) {
      const attemptDeadline = Math.min(deadline, Date.now() + 500);
      try {
        await connect(attemptDeadline);
        return requireStatus ? await status(attemptDeadline) : undefined;
      } catch (error) {
        if (client) { client.close(); client = undefined; }
        if (!retryable(error)) throw error;
        const remaining = deadline - Date.now();
        if (remaining > 0) await new Promise(resolve => setTimeout(resolve, Math.min(100, remaining)));
      }
    }
    throw deadlineError();
  };
  try {
    if (options.command === 'powerpoint launch') {
      let reply;
      const probeDeadline = Math.min(deadline, Date.now() + 500);
      try {
        await connect(probeDeadline);
        reply = await status(probeDeadline);
      } catch (error) {
        if (client) { client.close(); client = undefined; }
        if (!retryable(error)) throw error;
        if (Date.now() >= deadline) throw deadlineError();
        await launchPowerPoint(deadline);
        reply = await ready(true);
      }
      console.log(`PowerPoint ready (PID ${reply.processId}).`);
    } else {
      await ready(false);
      if (options.command === 'presentation new') {
        const reply = await client.request('PowerPoint.newPresentation', { title: options.title }, deadline);
        if (typeof reply.name !== 'string' || !Number.isSafeInteger(reply.slideCount) || reply.slideCount < 0) {
          throw new Error('Invalid server response: expected presentation name and slideCount. The presentation may already have been created; the command was not retried.');
        }
        console.log(`Created presentation ${JSON.stringify(reply.name)} (${reply.slideCount} slide${reply.slideCount === 1 ? '' : 's'}).`);
      } else {
        await client.request('PowerPoint.setSlideTitle', { slideIndex: options.slide, text: options.title }, deadline);
        console.log(`Updated slide ${options.slide} title to ${JSON.stringify(options.title)}.`);
      }
    }
  } catch (error) {
    const code = Number.isInteger(error.code) ? `Error ${error.code}: ` : '';
    const data = Object.hasOwn(error, 'data') ? `\nDetails: ${JSON.stringify(error.data)}` : '';
    const unavailable = retryable(error) || Date.now() >= deadline;
    const cancellation = error.code === -32002 || error.code === -32003
      ? '\nQueued work is cancelled; an Office call already started may complete. Do not blindly retry a command that changes a presentation.'
      : '';
    throw new Error(`${code}${error.message}${data}${cancellation}${unavailable ? '\nEnsure PowerPoint is running and the NetOffice native add-in is registered, enabled, and listening at ' + endpoint + '. --port must match the registered ServerPort; launch does not reconfigure the add-in. Use netoffice powerpoint launch on Windows or increase --timeout.' : ''}`);
  } finally {
    if (client) client.close();
  }
}

main().catch(error => {
  console.error(`netoffice: ${error.message}`);
  process.exitCode = 1;
});
