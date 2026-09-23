const fs = require('fs');
const path = require('path');
const {Writable} = require('stream');

const StreamBuf = verquire('utils/stream-buf');
const StringBuf = verquire('utils/string-buf');

describe('StreamBuf', () => {
  // StreamBuf is designed as a general-purpose writable-readable stream
  // However its use in ExcelJS is primarily as a memory buffer between
  // the streaming writers and the archive, hence the tests here will
  // focus just on that.
  it('writes strings as UTF8', () => {
    const stream = new StreamBuf();
    stream.write('Hello, World!');
    const chunk = stream.read();
    expect(chunk instanceof Buffer).to.be.ok();
    expect(chunk.toString('UTF8')).to.equal('Hello, World!');
  });

  it('writes StringBuf chunks', () => {
    const stream = new StreamBuf();
    const strBuf = new StringBuf({size: 64});
    strBuf.addText('Hello, World!');
    stream.write(strBuf);
    const chunk = stream.read();
    expect(chunk instanceof Buffer).to.be.ok();
    expect(chunk.toString('UTF8')).to.equal('Hello, World!');
  });

  it('signals end', done => {
    const stream = new StreamBuf();
    stream.on('finish', () => {
      done();
    });
    stream.write('Hello, World!');
    stream.end();
  });

  it('handles buffers', () =>
    new Promise((resolve, reject) => {
      const s = fs.createReadStream(path.join(__dirname, 'data/image1.png'));
      const sb = new StreamBuf();
      sb.on('finish', () => {
        const buf = sb.toBuffer();
        expect(buf.length).to.equal(1672);
        resolve();
      });
      sb.on('error', reject);
      s.pipe(sb);
    }));
  it('handle unsupported type of chunk', async () => {
    const stream = new StreamBuf();
    try {
      await stream.write({});
      expect.fail('should fail for given argument');
    } catch (e) {
      expect(e.message).to.equal(
        'Chunk must be one of type String, Buffer or StringBuf.'
      );
    }
  });

  it('honours pipe backpressure and resumes on drain', async () => {
    const stream = new StreamBuf({batch: true, bufSize: 8});
    const pendingCallbacks = [];
    const chunks = [];
    const dest = new Writable({
      highWaterMark: 1,
      write(chunk, encoding, callback) {
        chunks.push(chunk.toString());
        pendingCallbacks.push(callback);
      },
    });
    stream.pipe(dest);

    // first chunk fills the buffer, second chunk pushes the first into
    // the pipe, which immediately applies backpressure
    stream.write('AAAAAAAA');
    const p2 = stream.write('BBBBBBBB');
    await new Promise(resolve => setImmediate(resolve));
    expect(chunks).to.deep.equal(['AAAAAAAA']);

    let resolved = false;
    p2.then(() => {
      resolved = true;
    });
    await new Promise(resolve => setImmediate(resolve));
    expect(resolved).to.be.false();

    // releasing the pending pipe write triggers drain and the parked write
    pendingCallbacks.splice(0).forEach(callback => callback());
    await p2;
    expect(resolved).to.be.true();

    // after drain, more data can flow into the pipe (the pipe applies
    // backpressure again, so release its pending write too)
    const p4 = stream.write('CCCCCCCC');
    pendingCallbacks.splice(0).forEach(callback => callback());
    await p4;
    expect(chunks).to.deep.equal(['AAAAAAAA', 'BBBBBBBB']);
  });
});
