import { parentPort } from 'node:worker_threads';

parentPort.on('message', ({ request, reply }) => {
  if (request.document.sheets[0]?.cells[0]?.content.kind !== 'formula') {
    throw new Error('Expected the calculation fixture in the detached snapshot.');
  }
  parentPort.postMessage(reply);
});
