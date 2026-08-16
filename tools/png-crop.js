'use strict';
/* 裁剪截图顶部区域并另存（解码 + 重新编码 RGB PNG） */
var fs = require('fs');
var zlib = require('zlib');

var TABLE = [];
(function () {
  for (var n = 0; n < 256; n++) {
    var c = n;
    for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    TABLE[n] = c >>> 0;
  }
})();
function crc32(b) {
  var crc = 0xFFFFFFFF;
  for (var i = 0; i < b.length; i++) crc = TABLE[(crc ^ b[i]) & 0xFF] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}
function chunk(type, data) {
  var len = Buffer.alloc(4);
  len.writeUInt32BE(data.length, 0);
  var body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  var crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body), 0);
  return Buffer.concat([len, body, crc]);
}

var src = process.argv[2];
var dst = process.argv[3];
var y0 = Number(process.argv[4] || 0);
var y1 = Number(process.argv[5] || 120);

var buf = fs.readFileSync(src);
var pos = 8, width = 0, height = 0, colorType = 0;
var idat = [];
while (pos + 8 <= buf.length) {
  var len = buf.readUInt32BE(pos);
  var type = buf.toString('ascii', pos + 4, pos + 8);
  var data = buf.slice(pos + 8, pos + 8 + len);
  if (type === 'IHDR') {
    width = data.readUInt32BE(0);
    height = data.readUInt32BE(4);
    colorType = data[9];
  } else if (type === 'IDAT') idat.push(data);
  pos += 12 + len;
}
var channels = (colorType === 6) ? 4 : ((colorType === 2) ? 3 : 1);
var stride = width * channels;
var raw = zlib.inflateSync(Buffer.concat(idat));
var img = Buffer.alloc(stride * height);
var prev = Buffer.alloc(stride);
function paeth(a, b, c) {
  var p = a + b - c;
  var pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c);
  if (pa <= pb && pa <= pc) return a;
  if (pb <= pc) return b;
  return c;
}
for (var y = 0; y < height; y++) {
  var f = raw[y * (stride + 1)];
  var line = img.slice(y * stride, (y + 1) * stride);
  for (var i = 0; i < stride; i++) {
    var v = raw[y * (stride + 1) + 1 + i];
    var left = (i >= channels) ? line[i - channels] : 0;
    var up = prev[i];
    var ul = (i >= channels) ? prev[i - channels] : 0;
    if (f === 1) v = (v + left) & 255;
    else if (f === 2) v = (v + up) & 255;
    else if (f === 3) v = (v + ((left + up) >> 1)) & 255;
    else if (f === 4) v = (v + paeth(left, up, ul)) & 255;
    line[i] = v;
  }
  prev = line;
}

y0 = Math.max(0, Math.min(y0, height - 1));
y1 = Math.max(y0 + 1, Math.min(y1, height));
var newH = y1 - y0;
var scale = Number(process.argv[6] || 2);

/* 最近邻放大 scale 倍，帮助 OCR */
var outW = width * scale;
var outH = newH * scale;
var outStride = outW * 3;
var rawOut = Buffer.alloc(outH * (outStride + 1));
for (var yy = 0; yy < outH; yy++) {
  var sy = y0 + Math.floor(yy / scale);
  rawOut[yy * (outStride + 1)] = 0;
  for (var xx = 0; xx < outW; xx++) {
    var sx = Math.floor(xx / scale);
    var o = sy * stride + sx * channels;
    var oo = yy * (outStride + 1) + 1 + xx * 3;
    rawOut[oo] = img[o];
    rawOut[oo + 1] = img[o + 1];
    rawOut[oo + 2] = img[o + 2];
  }
}
var ihdr = Buffer.alloc(13);
ihdr.writeUInt32BE(outW, 0);
ihdr.writeUInt32BE(outH, 4);
ihdr[8] = 8;
ihdr[9] = 2;
var sig = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
var idatOut = zlib.deflateSync(rawOut, { level: 9 });
fs.writeFileSync(dst, Buffer.concat([sig, chunk('IHDR', ihdr), chunk('IDAT', idatOut), chunk('IEND', Buffer.alloc(0))]));
console.log('crop ' + src + ' -> ' + dst + ' y=' + y0 + '..' + y1 + ' out=' + outW + 'x' + outH);
