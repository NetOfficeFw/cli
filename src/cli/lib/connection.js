'use strict';

const WebSocket = require('ws');

const maximumMessageBytes = 1024 * 1024;

function deadlineError() {
  const error = new Error('The total operation deadline expired.');
  error.code = -32002;
  return error;
}

function disconnectedError(message) {
  const error = new Error(message);
  error.code = -32003;
  return error;
}

function retryable(error) {
  return error.code === -32002 || error.code === -32003;
}

function isObject(value) {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

class Connection {
  static async connect(url, deadline) {
    if (Date.now() >= deadline) throw deadlineError();
    const connection = new Connection(url, deadline);
    await connection.opened;
    return connection;
  }

  constructor(url, deadline) {
    this.pending = new Map();
    this.nextId = 1;
    this.failure = null;
    this.socket = new WebSocket(url, {
      maxPayload: maximumMessageBytes,
      perMessageDeflate: false,
      handshakeTimeout: Math.max(1, deadline - Date.now())
    });
    this.opened = new Promise((resolve, reject) => {
      this.rejectOpen = reject;
      this.openTimer = setTimeout(() => this.fail(deadlineError()), Math.max(1, deadline - Date.now()));
      this.socket.on('open', () => {
        clearTimeout(this.openTimer);
        if (!this.failure) resolve();
      });
    });
    this.socket.on('message', (data, binary) => this.receive(data, binary));
    this.socket.on('error', error => this.fail(disconnectedError(`WebSocket connection error: ${error.message}`)));
    this.socket.on('close', (code, reason) => this.fail(disconnectedError(`WebSocket connection closed (${code}${reason.length ? `: ${reason.toString()}` : ''}).`)));
  }

  request(method, params, deadline) {
    if (this.failure) return Promise.reject(this.failure);
    const remaining = deadline - Date.now();
    if (remaining <= 0) {
      this.fail(deadlineError());
      return Promise.reject(this.failure);
    }
    if (this.socket.readyState !== WebSocket.OPEN) return Promise.reject(disconnectedError('WebSocket connection is not open.'));
    if (!Number.isSafeInteger(this.nextId)) return Promise.reject(new Error('Request ID limit exceeded.'));
    const id = this.nextId++;
    const message = JSON.stringify({ id, method, params, timeoutMs: Math.min(remaining, 2147483647) });
    if (Buffer.byteLength(message, 'utf8') > maximumMessageBytes) return Promise.reject(new Error('The request exceeds the 1 MiB message limit.'));
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => this.fail(deadlineError()), remaining);
      this.pending.set(id, { resolve, reject, timer });
      try {
        this.socket.send(message, error => {
          if (error) this.fail(disconnectedError(`Unable to send the request: ${error.message}`));
        });
      } catch (error) {
        this.fail(disconnectedError(`Unable to send the request: ${error.message}`));
      }
    });
  }

  receive(data, binary) {
    if (this.failure) return;
    let reply;
    try {
      if (binary) throw new Error('Expected a JSON text message, not binary data.');
      reply = JSON.parse(data.toString('utf8'));
      if (!isObject(reply)) throw new Error('Expected a reply object.');
      const hasResult = Object.hasOwn(reply, 'result');
      const hasError = Object.hasOwn(reply, 'error');
      if (hasResult === hasError) throw new Error('Expected exactly one of result or error.');
      if (hasError && (!isObject(reply.error) || !Number.isInteger(reply.error.code) || typeof reply.error.message !== 'string')) {
        throw new Error('Invalid error envelope.');
      }
      if (hasResult && !isObject(reply.result)) throw new Error('Invalid result envelope.');
      if (reply.id === null && hasError) {
        const error = this.serverError(reply.error);
        this.fail(error);
        return;
      }
      if (!Number.isSafeInteger(reply.id) || reply.id < 0 || !this.pending.has(reply.id)) throw new Error('Invalid or unknown reply ID.');
    } catch (error) {
      this.fail(new Error(`Invalid server response: ${error.message}`));
      return;
    }
    const pending = this.pending.get(reply.id);
    this.pending.delete(reply.id);
    clearTimeout(pending.timer);
    if (Object.hasOwn(reply, 'error')) {
      const error = this.serverError(reply.error);
      pending.reject(error);
      if (error.code === -32002 || error.code === -32003) this.fail(error);
    } else pending.resolve(reply.result);
  }

  serverError(value) {
    const error = new Error(value.message);
    error.code = value.code;
    if (Object.hasOwn(value, 'data')) error.data = value.data;
    return error;
  }

  fail(error) {
    if (this.failure) return;
    this.failure = error;
    clearTimeout(this.openTimer);
    this.rejectOpen(error);
    for (const pending of this.pending.values()) {
      clearTimeout(pending.timer);
      pending.reject(error);
    }
    this.pending.clear();
    // Abort the socket immediately so the native session cancels queued STA work.
    if (this.socket.readyState !== WebSocket.CLOSED) this.socket.terminate();
  }

  close() {
    this.fail(disconnectedError('WebSocket session closed.'));
  }
}

module.exports = { Connection, deadlineError, retryable };
