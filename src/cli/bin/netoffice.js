#!/usr/bin/env node
'use strict';

const { parseArguments, usage } = require('../lib/arguments');
const fs = require('node:fs');
const path = require('node:path');
const { execFile } = require('node:child_process');

// Fixed script: no user-supplied text is evaluated by PowerShell.
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

// The add-in acknowledges the unsaved-document preflight before this separate
// process requests Quit; quitting from inside its own STA callback stalls Office.
const shutdownScript = `
$ErrorActionPreference = 'Stop'
$running = @(Get-Process POWERPNT)
if ($running.Count -ne 1 -or $running[0].Id -ne [int]$env:NETOFFICE_EXPECTED_PID) {
  throw 'PowerPoint process identity is ambiguous or differs from the add-in process.'
}
$app = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
foreach ($presentation in @($app.Presentations)) {
  if ($env:NETOFFICE_FORCE -eq '1') {
    $presentation.Saved = -1
    $presentation.Close()
  } elseif ($presentation.Saved -ne -1 -or [string]::IsNullOrEmpty($presentation.Path)) {
    throw 'An open presentation has unsaved changes or no saved path.'
  }
}
$app.Quit()
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

function quitPowerPoint(processId, deadline, force) {
  return new Promise((resolve, reject) => {
    const executable = path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
    execFile(executable, ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', shutdownScript], {
      windowsHide: true,
      timeout: Math.max(1, deadline - Date.now()),
      maxBuffer: 1024 * 1024,
      env: { ...process.env, NETOFFICE_EXPECTED_PID: String(processId), NETOFFICE_FORCE: force ? '1' : '0' }
    }, (error, stdout, stderr) => {
      if (error) reject(new Error(`Unable to shut down PowerPoint: ${stderr.trim() || error.message}`));
      else resolve();
    });
  });
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  if (options.help) { process.stdout.write(usage); return; }
  if (options.command.startsWith('powerpoint ') && process.platform !== 'win32') {
    throw new Error(`${options.command} is supported only on Windows with desktop Microsoft PowerPoint installed.`);
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
  // One WebSocket request for a parsed command; mutations are never retried.
  const send = async parsed => {
    const params = parsed.command === 'presentation open' ? { path: path.resolve(parsed.path) } : parsed.params;
    const reply = await client.request(parsed.method, params, deadline);
    if (parsed.command === 'presentation new' && (typeof reply.name !== 'string' ||
        !Number.isSafeInteger(reply.slideCount) || reply.slideCount !== 0 ||
        typeof reply.url !== 'string' || typeof reply.id !== 'string')) {
      throw new Error('Invalid server response: presentation may already have been created; command was not retried.');
    }
    return reply;
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
      if (options.command === 'powerpoint shutdown') {
        const { processId } = await status(deadline);
        await client.request(options.method, options.params, deadline);
        client.close();
        client = undefined;
        await quitPowerPoint(processId, deadline, options.force === true);
        while (Date.now() < deadline) {
          try {
            process.kill(processId, 0);
          } catch (error) {
            if (error.code === 'ESRCH') {
              console.log(JSON.stringify({ processId, stopped: true }));
              return;
            }
            throw error;
          }
          await new Promise(resolve => setTimeout(resolve, Math.min(100, deadline - Date.now())));
        }
        throw new Error(`PowerPoint (PID ${processId}) did not exit before the deadline.`);
      } else if (options.command === 'presentation list') {
        const response = await fetch(`http://127.0.0.1:${options.port}/json/list`, {
          signal: AbortSignal.timeout(Math.max(1, deadline - Date.now()))
        });
        if (!response.ok) throw new Error(`Unable to list presentations (HTTP ${response.status}).`);
        console.log(JSON.stringify(await response.json()));
      } else if (options.command === 'batch run') {
        const steps = JSON.parse(fs.readFileSync(path.resolve(options.file), 'utf8'));
        if (!Array.isArray(steps) || steps.some(step => !Array.isArray(step) || step.some(arg => typeof arg !== 'string'))) {
          throw new Error('batch run --file requires a JSON array of argument arrays.');
        }
        const results = [];
        for (const [index, step] of steps.entries()) {
          const args = step.map(arg => arg.replace(/\$(\d+)\.([A-Za-z][A-Za-z0-9]*)/g, (reference, position, field) => {
            const source = results[Number(position)];
            if (!source || !Object.hasOwn(source, field)) {
              throw new Error(`Batch command ${index}: ${reference} does not name a field of an earlier result.`);
            }
            return String(source[field]);
          }));
          const stepOptions = parseArguments(args);
          if (!stepOptions.method || stepOptions.command.startsWith('powerpoint ')) {
            throw new Error(`Batch command ${index}: ${stepOptions.command || args.join(' ')} cannot run in a batch.`);
          }
          try {
            results.push(await send(stepOptions));
          } catch (error) {
            console.log(JSON.stringify(results));
            error.message = `Batch command ${index} (${stepOptions.command}) failed after ${index} completed: ${error.message}`;
            throw error;
          }
        }
        console.log(JSON.stringify(results));
      } else {
        console.log(JSON.stringify(await send(options)));
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
