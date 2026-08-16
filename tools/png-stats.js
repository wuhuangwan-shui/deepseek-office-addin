'use strict';
/* 极简 PNG 解码器：解码截图并统计暗像素分布（诊断截图是否空白） */
var fs = require('fs');
var zlib = require('zlib');

var file = process.argv[2];
var buf = fs.readFileSync(file);
var pos = 8;
var width = 0, height = 0, bitDepth = 0, colorType = 0;
var idat = [];
while (pos + 8 <= buf.length) {
  var len = buf.readUInt32BE(pos);
  var type = buf.toString('ascii', pos + 4, pos + 8);
  var data = buf.slice(pos + 8, pos + 8 + len);
  if (type === 'IHDR') {
    width = data.readUInt32BE(0);
    height = data.readUInt32BE(4);
    bitDepth = data[8];
    colorType = data[9];
  } else if (type === 'IDAT') {
    idat.push(data);
  }
  pos += 12 + len;
}
var channels = (colorType === 6) ? 4 : ((colorType === 2) ? 3 : ((colorType === 4) ? 2 : 1));
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
  var filter = raw[y * (stride + 1)];
  var line = img.slice(y * stride, (y + 1) * stride);
  for (var i = 0; i < stride; i++) {
    var v = raw[y * (stride + 1) + 1 + i];
    var left = (i >= channels) ? line[i - channels] : 0;
    var up = prev[i];
    var ul = (i >= channels) ? prev[i - channels] : 0;
    if (filter === 1) v = (v + left) & 255;
    else if (filter === 2) v = (v + up) & 255;
    else if (filter === 3) v = (v + ((left + up) >> 1)) & 255;
    else if (filter === 4) v = (v + paeth(left, up, ul)) & 255;
    line[i] = v;
  }
  prev = line;
}

var dark = 0;
var minX = width, maxX = -1, minY = height, maxY = -1;
for (var y = 0; y < height; y++) {
  for (var x = 0; x < width; x++) {
    var o = y * stride + x * channels;
    var lum = 0.299 * img[o] + 0.587 * img[o + 1] + 0.114 * img[o + 2];
    if (lum < 150) {
      dark++;
      if (x < minX) minX = x;
      if (x > maxX) maxX = x;
      if (y < minY) minY = y;
      if (y > maxY) maxY = y;
    }
  }
}
console.log('size=' + width + 'x' + height + ' colorType=' + colorType + ' bitDepth=' + bitDepth);
console.log('darkPixels=' + dark + ' darkRatio=' + (dark / (width * height)).toFixed(4));
console.log('darkBox=(' + minX + ',' + minY + ')-(' + maxX + ',' + maxY + ')');

var cols = 32, rows = 4;
var cw = Math.ceil(width / cols), ch = Math.ceil(height / rows);
var NL = String.fromCharCode(10);
var out = '';
for (var ry = 0; ry < rows; ry++) {
  var lineStr = '';
  for (var rx = 0; rx < cols; rx++) {
    var cnt = 0, d = 0;
    for (var yy = ry * ch; yy < Math.min((ry + 1) * ch, height); yy += 3) {
      for (var xx = rx * cw; xx < Math.min((rx + 1) * cw, width); xx += 3) {
        var oo = yy * stride + xx * channels;
        var lum2 = 0.299 * img[oo] + 0.587 * img[oo + 1] + 0.114 * img[oo + 2];
        if (lum2 < 200) d++;
        cnt++;
      }
    }
    var ratio = cnt ? (d / cnt) : 0;
    lineStr += (ratio > 0.3) ? '#' : ((ratio > 0.1) ? '+' : ((ratio > 0.02) ? '.' : ' '));
  }
  out += lineStr + NL;
}
console.log('map:' + NL + out);
