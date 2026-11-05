/**
 * pdfjs-dist 5.x engines say Node >=20. Electron 17 runs Node 16.
 * Legacy pdf.mjs 5.5.207 has no private-fields/`using` syntax — it fails on
 * missing globals. Polyfill those only; keep the 5.5 extractor.
 */

import { createRequire } from 'node:module';
import {
  brotliCompressSync,
  brotliDecompressSync,
  deflateRawSync,
  deflateSync,
  gunzipSync,
  gzipSync,
  inflateRawSync,
  inflateSync,
} from 'node:zlib';

const nodeRequire = createRequire(process.execPath);

export function ensurePdfRuntimePolyfills(): void {
  const g = globalThis as typeof globalThis & {
    structuredClone?: (value: unknown) => unknown;
  };

  const PromiseCtor = Promise as typeof Promise & {
    withResolvers?: <T>() => {
      promise: Promise<T>;
      resolve: (value: T | PromiseLike<T>) => void;
      reject: (reason?: unknown) => void;
    };
  };
  if (typeof PromiseCtor.withResolvers !== 'function') {
    PromiseCtor.withResolvers = function withResolvers<T>() {
      let resolve!: (value: T | PromiseLike<T>) => void;
      let reject!: (reason?: unknown) => void;
      const promise = new Promise<T>((res, rej) => {
        resolve = res;
        reject = rej;
      });
      return { promise, resolve, reject };
    };
  }

  if (typeof Object.hasOwn !== 'function') {
    Object.defineProperty(Object, 'hasOwn', {
      value(obj: object, prop: PropertyKey) {
        return Object.prototype.hasOwnProperty.call(obj, prop);
      },
      configurable: true,
      writable: true,
    });
  }

  if (typeof Array.prototype.at !== 'function') {
    Object.defineProperty(Array.prototype, 'at', {
      value(this: unknown[], index: number) {
        const n = Math.trunc(index) || 0;
        const k = n < 0 ? this.length + n : n;
        return this[k];
      },
      configurable: true,
      writable: true,
    });
  }

  if (typeof g.structuredClone !== 'function') {
    g.structuredClone = (value: unknown) => cloneValue(value, new WeakMap());
  }

  // pdfjs does `new Response(stream).arrayBuffer()`.
  // Electron 17 has no Fetch globals. Node 20+ ships native Response, but
  // arrayBuffer() on the streams pdfjs uses for file:// cmaps/fonts can hang
  // long enough to trip the default 5s test timeout.
  const fetchGlobals = globalThis as { Response?: unknown; Headers?: unknown };
  fetchGlobals.Response = ResponsePoly;
  if (typeof fetchGlobals.Headers !== 'function') {
    fetchGlobals.Headers = HeadersPoly;
  }

  // pdf.mjs core-js does `NativeDOMException.prototype` while the module
  // evaluates. Electron 17 has no DOMException, so the import throws before
  // any page is read.
  ensureDomException();
  ensureWebStreams();
  // pdf.mjs evaluates `new DOMMatrix()` while the module loads.
  // Electron 17 has no canvas globals. Provide them, and the Node 20
  // helpers pdfjs uses to look up `module` / `url` / `@napi-rs/canvas`.
  ensureCanvasGlobals();
  ensureNodeBuiltins();
  ensureCanvasPackageStub();
}

function cloneValue(value: unknown, seen: WeakMap<object, unknown>): unknown {
  if (value === null || typeof value !== 'object') return value;
  if (value instanceof Date) return new Date(value.getTime());
  if (value instanceof RegExp) return new RegExp(value.source, value.flags);
  if (value instanceof ArrayBuffer) return value.slice(0);
  if (ArrayBuffer.isView(value)) {
    const bytes = new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
    return new Uint8Array(bytes);
  }
  const obj = value as object;
  if (seen.has(obj)) return seen.get(obj);
  if (Array.isArray(value)) {
    const arr: unknown[] = [];
    seen.set(obj, arr);
    for (const item of value) arr.push(cloneValue(item, seen));
    return arr;
  }
  const out: Record<string, unknown> = {};
  seen.set(obj, out);
  for (const key of Object.keys(value as Record<string, unknown>)) {
    out[key] = cloneValue((value as Record<string, unknown>)[key], seen);
  }
  return out;
}

class HeadersPoly {
  private readonly map = new Map<string, string>();

  constructor(init?: Record<string, string> | Array<[string, string]>) {
    if (!init) return;
    const entries = Array.isArray(init) ? init : Object.entries(init);
    for (const [key, value] of entries) this.map.set(key.toLowerCase(), String(value));
  }

  get(name: string): string | null {
    return this.map.get(name.toLowerCase()) ?? null;
  }

  has(name: string): boolean {
    return this.map.has(name.toLowerCase());
  }

  set(name: string, value: string): void {
    this.map.set(name.toLowerCase(), String(value));
  }
}

class ResponsePoly {
  readonly status: number;
  readonly ok: boolean;
  readonly headers: HeadersPoly;
  readonly body: unknown;

  constructor(
    body?: unknown,
    init?: { status?: number; headers?: Record<string, string> | Array<[string, string]> }
  ) {
    this.body = body ?? null;
    this.status = init?.status ?? 200;
    this.ok = this.status >= 200 && this.status < 300;
    this.headers = new HeadersPoly(init?.headers);
  }

  async arrayBuffer(): Promise<ArrayBuffer> {
    return bodyToArrayBuffer(this.body);
  }

  async text(): Promise<string> {
    return new TextDecoder().decode(await this.arrayBuffer());
  }

  async json(): Promise<unknown> {
    return JSON.parse(await this.text());
  }
}

async function bodyToArrayBuffer(body: unknown): Promise<ArrayBuffer> {
  if (body == null) return new ArrayBuffer(0);
  if (body instanceof ArrayBuffer) return body;
  if (ArrayBuffer.isView(body)) {
    const view = body as ArrayBufferView;
    const copy = new Uint8Array(view.byteLength);
    copy.set(new Uint8Array(view.buffer, view.byteOffset, view.byteLength));
    return copy.buffer;
  }
  if (typeof body === 'string') return new TextEncoder().encode(body).buffer;
  if (typeof Blob !== 'undefined' && body instanceof Blob) return body.arrayBuffer();

  const web = body as { getReader?: () => { read: () => Promise<{ done: boolean; value?: Uint8Array }> } };
  if (typeof web.getReader === 'function') {
    const reader = web.getReader();
    const chunks: Uint8Array[] = [];
    let total = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      if (!value) continue;
      const bytes = value instanceof Uint8Array ? value : new Uint8Array(value);
      chunks.push(bytes);
      total += bytes.byteLength;
    }
    return concatBytes(chunks, total);
  }

  const iterable = body as AsyncIterable<Uint8Array>;
  if (iterable && typeof iterable[Symbol.asyncIterator] === 'function') {
    const chunks: Uint8Array[] = [];
    let total = 0;
    for await (const chunk of iterable) {
      const bytes = chunk instanceof Uint8Array ? chunk : new Uint8Array(chunk);
      chunks.push(bytes);
      total += bytes.byteLength;
    }
    return concatBytes(chunks, total);
  }

  throw new TypeError('Unsupported Response body');
}

function concatBytes(chunks: Uint8Array[], total: number): ArrayBuffer {
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return out.buffer;
}

const DOM_EXCEPTION_CODES: Record<string, number> = {
  IndexSizeError: 1,
  HierarchyRequestError: 3,
  WrongDocumentError: 4,
  InvalidCharacterError: 5,
  NoModificationAllowedError: 7,
  NotFoundError: 8,
  NotSupportedError: 9,
  InUseAttributeError: 10,
  InvalidStateError: 11,
  SyntaxError: 12,
  InvalidModificationError: 13,
  NamespaceError: 14,
  InvalidAccessError: 15,
  TypeMismatchError: 17,
  SecurityError: 18,
  NetworkError: 19,
  AbortError: 20,
  URLMismatchError: 21,
  QuotaExceededError: 22,
  TimeoutError: 23,
  InvalidNodeTypeError: 24,
  DataCloneError: 25,
};

class DOMExceptionPoly extends Error {
  readonly code: number;

  constructor(message = '', name = 'Error') {
    super(String(message));
    this.name = String(name);
    this.code = DOM_EXCEPTION_CODES[this.name] ?? 0;
  }
}

function ensureDomException(): void {
  const g = globalThis as { DOMException?: unknown };
  if (typeof g.DOMException === 'function') return;
  for (const [name, code] of Object.entries(DOM_EXCEPTION_CODES)) {
    Object.defineProperty(DOMExceptionPoly, name, { value: code });
    Object.defineProperty(DOMExceptionPoly.prototype, name, { value: code });
  }
  g.DOMException = DOMExceptionPoly;
}

type ReadResult = { done: boolean; value?: Uint8Array };

type ByteReader = {
  read(): Promise<ReadResult>;
  releaseLock(): void;
};

type StreamController = {
  enqueue(chunk: Uint8Array): void;
  close(): void;
  error(err?: unknown): void;
};

type WebStreamModule = {
  ReadableStream: {
    new (source?: { start?: (controller: StreamController) => void }): {
      getReader(): ByteReader;
    };
    prototype: {
      getReader(): ByteReader;
      [Symbol.asyncIterator]?: () => AsyncIterator<Uint8Array>;
    };
  };
  WritableStream: new (sink?: {
    write?: (chunk: Uint8Array) => void | Promise<void>;
    close?: () => void | Promise<void>;
    abort?: (reason?: unknown) => void | Promise<void>;
  }) => unknown;
  TransformStream: new (...args: unknown[]) => unknown;
};

function withoutExperimentalStreamWarning<T>(load: () => T): T {
  const orig = process.emitWarning;
  process.emitWarning = ((warning: unknown, type?: unknown, code?: unknown, ...rest: unknown[]) => {
    const text =
      typeof warning === 'string'
        ? warning
        : warning && typeof warning === 'object' && 'message' in warning
          ? String((warning as { message?: unknown }).message ?? '')
          : '';
    const name =
      warning && typeof warning === 'object' && 'name' in warning
        ? String((warning as { name?: unknown }).name ?? '')
        : '';
    if (
      type === 'ExperimentalWarning' ||
      code === 'ExperimentalWarning' ||
      name === 'ExperimentalWarning' ||
      text.includes('stream/web')
    ) {
      return;
    }
    const args: unknown[] = [warning];
    if (type !== undefined) args.push(type);
    if (code !== undefined) args.push(code);
    (orig as { apply(thisArg: unknown, args: unknown[]): void }).apply(process, args);
  }) as typeof process.emitWarning;
  try {
    return load();
  } finally {
    process.emitWarning = orig;
  }
}

function webStreams(): WebStreamModule {
  return withoutExperimentalStreamWarning(
    () => nodeRequire('node:stream/web') as WebStreamModule
  );
}

function ensureNodeBuiltins(): void {
  const proc = process as NodeJS.Process & {
    getBuiltinModule?: (name: string) => unknown;
  };
  if (typeof proc.getBuiltinModule === 'function') return;
  proc.getBuiltinModule = (name: string) => {
    const id = name.startsWith('node:') ? name : `node:${name}`;
    return nodeRequire(id);
  };
}

type NodeModuleCtor = {
  prototype: {
    require: (this: unknown, id: string) => unknown;
  };
};

function ensureCanvasPackageStub(): void {
  const Module = nodeRequire('node:module') as NodeModuleCtor;
  const orig = Module.prototype.require;
  if ((orig as { huginCanvasStub?: boolean }).huginCanvasStub) return;
  function patched(this: unknown, id: string): unknown {
    if (id === '@napi-rs/canvas') {
      try {
        return orig.call(this, id);
      } catch {
        const g = globalThis as {
          DOMMatrix?: unknown;
          ImageData?: unknown;
          Path2D?: unknown;
        };
        return {
          DOMMatrix: g.DOMMatrix,
          ImageData: g.ImageData,
          Path2D: g.Path2D,
        };
      }
    }
    return orig.call(this, id);
  }
  (patched as { huginCanvasStub?: boolean }).huginCanvasStub = true;
  Module.prototype.require = patched;
}

function ensureAsyncIterator(Ctor: WebStreamModule['ReadableStream']): void {
  const proto = Ctor.prototype;
  if (typeof proto[Symbol.asyncIterator] === 'function') return;
  proto[Symbol.asyncIterator] = async function* (
    this: { getReader(): ByteReader }
  ): AsyncGenerator<Uint8Array> {
    const reader = this.getReader();
    try {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) return;
        if (value) yield value;
      }
    } finally {
      reader.releaseLock();
    }
  };
}

function concatChunks(chunks: Uint8Array[]): Buffer {
  const total = chunks.reduce((sum, chunk) => sum + chunk.byteLength, 0);
  const out = Buffer.allocUnsafe(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return out;
}

function inflateByName(format: string, input: Buffer): Buffer {
  switch (format) {
    case 'deflate':
      return inflateSync(input);
    case 'deflate-raw':
      return inflateRawSync(input);
    case 'gzip':
      return gunzipSync(input);
    case 'brotli':
      return brotliDecompressSync(input);
    default:
      throw new TypeError(`Unsupported decompression format: ${format}`);
  }
}

function deflateByName(format: string, input: Buffer): Buffer {
  switch (format) {
    case 'deflate':
      return deflateSync(input);
    case 'deflate-raw':
      return deflateRawSync(input);
    case 'gzip':
      return gzipSync(input);
    case 'brotli':
      return brotliCompressSync(input);
    default:
      throw new TypeError(`Unsupported compression format: ${format}`);
  }
}

class CodecStream {
  readonly readable: unknown;
  readonly writable: unknown;

  constructor(format: string, decode: boolean) {
    const { ReadableStream, WritableStream } = webStreams();
    const chunks: Uint8Array[] = [];
    let controller!: StreamController;
    this.readable = new ReadableStream({
      start(controllerIn) {
        controller = controllerIn;
      },
    });
    this.writable = new WritableStream({
      write(chunk) {
        chunks.push(chunk instanceof Uint8Array ? chunk : new Uint8Array(chunk));
      },
      close() {
        const output = decode
          ? inflateByName(format, concatChunks(chunks))
          : deflateByName(format, concatChunks(chunks));
        controller.enqueue(new Uint8Array(output));
        controller.close();
      },
      abort(reason) {
        controller.error(reason);
      },
    });
  }
}

function ensureWebStreams(): void {
  const g = globalThis as {
    ReadableStream?: unknown;
    WritableStream?: unknown;
    TransformStream?: unknown;
    CompressionStream?: unknown;
    DecompressionStream?: unknown;
  };
  const streams = webStreams();
  if (typeof g.ReadableStream !== 'function') {
    ensureAsyncIterator(streams.ReadableStream);
    g.ReadableStream = streams.ReadableStream;
  } else {
    ensureAsyncIterator(g.ReadableStream as WebStreamModule['ReadableStream']);
  }
  if (typeof g.WritableStream !== 'function') g.WritableStream = streams.WritableStream;
  if (typeof g.TransformStream !== 'function') g.TransformStream = streams.TransformStream;
  if (typeof g.DecompressionStream !== 'function') {
    g.DecompressionStream = class DecompressionStream {
      readonly readable: unknown;
      readonly writable: unknown;
      constructor(format: string) {
        const codec = new CodecStream(format, true);
        this.readable = codec.readable;
        this.writable = codec.writable;
      }
    };
  }
  if (typeof g.CompressionStream !== 'function') {
    g.CompressionStream = class CompressionStream {
      readonly readable: unknown;
      readonly writable: unknown;
      constructor(format: string) {
        const codec = new CodecStream(format, false);
        this.readable = codec.readable;
        this.writable = codec.writable;
      }
    };
  }
}

class DOMPointPoly {
  x: number;
  y: number;
  z: number;
  w: number;

  constructor(x = 0, y = 0, z = 0, w = 1) {
    this.x = x;
    this.y = y;
    this.z = z;
    this.w = w;
  }
}

class DOMMatrixPoly {
  a = 1;
  b = 0;
  c = 0;
  d = 1;
  e = 0;
  f = 0;
  m11 = 1;
  m12 = 0;
  m13 = 0;
  m14 = 0;
  m21 = 0;
  m22 = 1;
  m23 = 0;
  m24 = 0;
  m31 = 0;
  m32 = 0;
  m33 = 1;
  m34 = 0;
  m41 = 0;
  m42 = 0;
  m43 = 0;
  m44 = 1;
  is2D = true;

  constructor(init?: unknown) {
    if (typeof init === 'string') {
      const m = /matrix\(\s*([^)]+)\)/.exec(init);
      if (m) this.setComponents(m[1].split(/[,\s]+/).map(Number));
      return;
    }
    if (Array.isArray(init) || ArrayBuffer.isView(init)) {
      this.setComponents(Array.from(init as ArrayLike<number>));
      return;
    }
    if (init && typeof init === 'object' && 'a' in init) {
      const other = init as DOMMatrixPoly;
      this.setComponents([other.a, other.b, other.c, other.d, other.e, other.f]);
    }
  }

  get isIdentity(): boolean {
    return (
      this.a === 1 &&
      this.b === 0 &&
      this.c === 0 &&
      this.d === 1 &&
      this.e === 0 &&
      this.f === 0
    );
  }

  multiplySelf(other: DOMMatrixPoly): this {
    return this.setComponents([
      this.a * other.a + this.c * other.b,
      this.b * other.a + this.d * other.b,
      this.a * other.c + this.c * other.d,
      this.b * other.c + this.d * other.d,
      this.a * other.e + this.c * other.f + this.e,
      this.b * other.e + this.d * other.f + this.f,
    ]);
  }

  preMultiplySelf(other: DOMMatrixPoly): this {
    const next = new DOMMatrixPoly([
      other.a,
      other.b,
      other.c,
      other.d,
      other.e,
      other.f,
    ]);
    next.multiplySelf(this);
    return this.setComponents([next.a, next.b, next.c, next.d, next.e, next.f]);
  }

  multiply(other: DOMMatrixPoly): DOMMatrixPoly {
    return new DOMMatrixPoly(this).multiplySelf(other);
  }

  translateSelf(tx = 0, ty = 0): this {
    return this.multiplySelf(new DOMMatrixPoly([1, 0, 0, 1, tx, ty]));
  }

  translate(tx = 0, ty = 0): DOMMatrixPoly {
    return new DOMMatrixPoly(this).translateSelf(tx, ty);
  }

  scaleSelf(sx = 1, sy = sx): this {
    return this.multiplySelf(new DOMMatrixPoly([sx, 0, 0, sy, 0, 0]));
  }

  scale(sx = 1, sy = sx): DOMMatrixPoly {
    return new DOMMatrixPoly(this).scaleSelf(sx, sy);
  }

  invertSelf(): this {
    const det = this.a * this.d - this.b * this.c;
    if (!det) return this.setComponents([NaN, NaN, NaN, NaN, NaN, NaN]);
    return this.setComponents([
      this.d / det,
      -this.b / det,
      -this.c / det,
      this.a / det,
      (this.c * this.f - this.d * this.e) / det,
      (this.b * this.e - this.a * this.f) / det,
    ]);
  }

  inverse(): DOMMatrixPoly {
    return new DOMMatrixPoly(this).invertSelf();
  }

  transformPoint(point?: { x?: number; y?: number; z?: number; w?: number }): DOMPointPoly {
    const x = point?.x ?? 0;
    const y = point?.y ?? 0;
    return new DOMPointPoly(this.a * x + this.c * y + this.e, this.b * x + this.d * y + this.f);
  }

  toFloat32Array(): Float32Array {
    return Float32Array.from([this.a, this.b, this.c, this.d, this.e, this.f]);
  }

  toFloat64Array(): Float64Array {
    return Float64Array.from([this.a, this.b, this.c, this.d, this.e, this.f]);
  }

  private setComponents(values: number[]): this {
    const [a, b, c, d, e, f] = values;
    this.a = a ?? 1;
    this.b = b ?? 0;
    this.c = c ?? 0;
    this.d = d ?? 1;
    this.e = e ?? 0;
    this.f = f ?? 0;
    this.m11 = this.a;
    this.m12 = this.b;
    this.m21 = this.c;
    this.m22 = this.d;
    this.m41 = this.e;
    this.m42 = this.f;
    return this;
  }
}

class ImageDataPoly {
  readonly data: Uint8ClampedArray;
  readonly width: number;
  readonly height: number;

  constructor(dataOrWidth: Uint8ClampedArray | number, widthOrHeight: number, height?: number) {
    if (typeof dataOrWidth === 'number') {
      this.width = dataOrWidth;
      this.height = widthOrHeight;
      this.data = new Uint8ClampedArray(this.width * this.height * 4);
      return;
    }
    this.data = dataOrWidth;
    this.width = widthOrHeight;
    this.height = height ?? dataOrWidth.length / (widthOrHeight * 4);
  }
}

class Path2DPoly {
  constructor(_path?: unknown) {}
  addPath(): void {}
  closePath(): void {}
  moveTo(): void {}
  lineTo(): void {}
  bezierCurveTo(): void {}
  quadraticCurveTo(): void {}
  arc(): void {}
  arcTo(): void {}
  ellipse(): void {}
  rect(): void {}
}

function ensureCanvasGlobals(): void {
  const g = globalThis as {
    DOMMatrix?: unknown;
    DOMPoint?: unknown;
    ImageData?: unknown;
    Path2D?: unknown;
  };
  if (typeof g.DOMPoint !== 'function') g.DOMPoint = DOMPointPoly;
  if (typeof g.DOMMatrix !== 'function') g.DOMMatrix = DOMMatrixPoly;
  if (typeof g.ImageData !== 'function') g.ImageData = ImageDataPoly;
  if (typeof g.Path2D !== 'function') g.Path2D = Path2DPoly;
}
