// Development-only, localhost-only QA. Never distributed with the theme assets.
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { realpath } from 'node:fs/promises';

export const devRoot = path.resolve(process.env.OBSIDIAN_DEV_ROOT ?? path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../../..'));
const expectedVault = path.join(devRoot, '@Nexus-dev');
export async function connect() {
  if (await realpath(expectedVault) !== expectedVault) throw Error('Unexpected QA vault link.');
  const endpoint = process.env.OBSIDIAN_CDP ?? 'http://127.0.0.1:9235';
  const url = new URL(endpoint);
  if (!['127.0.0.1', 'localhost'].includes(url.hostname)) throw Error('Local CDP only');
  const targets = await (await fetch(endpoint + '/json/list')).json();
  for (const target of targets.filter(item => item.type === 'page' && item.url.includes('obsidian.md'))) {
    const socket = new WebSocket(target.webSocketDebuggerUrl);
    await new Promise((resolve, reject) => { socket.addEventListener('open', resolve, { once: true }); socket.addEventListener('error', reject, { once: true }); });
    let id = 0;
    const pending = new Map();
    socket.addEventListener('message', ({ data }) => {
      const message = JSON.parse(data);
      const request = pending.get(message.id);
      if (request) { pending.delete(message.id); clearTimeout(request.timer); message.error ? request.reject(Error(JSON.stringify(message.error))) : request.resolve(message.result); }
    });
    const send = (method, params = {}) => new Promise((resolve, reject) => {
      const requestId = ++id;
      const timer = setTimeout(() => { pending.delete(requestId); reject(Error('CDP timeout: ' + method)); }, 20000);
      pending.set(requestId, { resolve, reject, timer });
      socket.send(JSON.stringify({ id: requestId, method, params }));
    });
    const evaluate = async expression => {
      const result = await send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true });
      if (result.exceptionDetails) throw Error(result.exceptionDetails.text + ': ' + JSON.stringify(result.exceptionDetails.exception));
      return result.result?.value;
    };
    const vault = await evaluate('globalThis.app?.vault?.adapter?.basePath?.replaceAll("\\\\", "/")');
    if (vault === expectedVault.replaceAll('\\', '/')) return { send, evaluate, close: () => socket.close() };
    socket.close();
  }
  throw Error('No canonical @Nexus-dev Obsidian target found.');
}
