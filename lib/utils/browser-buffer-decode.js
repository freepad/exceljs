// eslint-disable-next-line node/no-unsupported-features/node-builtins
const TextDecoderCtor = typeof TextDecoder === 'undefined' ? null : TextDecoder;

function bufferToString(chunk) {
  if (typeof chunk === 'string') {
    return chunk;
  }
  return chunk.toString('utf8');
}

// Stream chunks can split a multi-byte UTF-8 character. `stream: true` buffers the
// trailing bytes so the next chunk completes them instead of yielding U+FFFD.
// The pending bytes are state, so one decoder must not be shared between parsers.
function createStreamDecoder() {
  const decoder = TextDecoderCtor ? new TextDecoderCtor('utf-8') : null;
  return {
    decode(chunk) {
      if (!decoder) {
        return typeof chunk === 'string' ? chunk : chunk.toString('utf8');
      }
      if (typeof chunk === 'string') {
        return decoder.decode() + chunk;
      }
      return decoder.decode(chunk, {stream: true});
    },
    flush() {
      return decoder ? decoder.decode() : '';
    },
  };
}

exports.bufferToString = bufferToString;
exports.createStreamDecoder = createStreamDecoder;
