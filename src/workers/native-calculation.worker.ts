import { workbookProtocolVersion, type NativeWorkerReply, type NativeWorkerRequest } from '../application/workbook-protocol';
import { runNativeGeneration } from '../calculation/native-generation';

type Generation = ReturnType<typeof runNativeGeneration>['generation'];
let accepted: { requestId: string; sessionId: string; generation: Generation } | undefined;
let pending: typeof accepted;

self.onmessage = ({ data }: MessageEvent<NativeWorkerRequest>) => {
  const { document, scope, acceptedRequestId, locale, kind, ...basis } = data;
  if (kind !== 'calculate-native' || basis.protocolVersion !== workbookProtocolVersion) return;
  if (pending?.requestId === acceptedRequestId && pending?.sessionId === basis.sessionId) accepted = pending;
  const previous = accepted?.requestId === acceptedRequestId && accepted?.sessionId === basis.sessionId
    ? accepted.generation : undefined;
  const { generation, updates, staleCellIds } = runNativeGeneration(document, scope, previous, { locale });
  pending = { requestId: basis.requestId, sessionId: basis.sessionId, generation };
  self.postMessage({ ...basis, kind: 'native-calculated', updates, staleCellIds } satisfies NativeWorkerReply);
};
