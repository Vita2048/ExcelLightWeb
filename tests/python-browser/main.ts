import { PythonProbe } from '../../src/python/probe-client';

declare global {
  interface Window { pythonProbe: PythonProbe }
}

window.pythonProbe = new PythonProbe();
