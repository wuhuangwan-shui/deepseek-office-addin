'use strict';
var urls = process.argv.slice(2);
function go(i) {
  if (i >= urls.length) return;
  var u = urls[i];
  console.log('===== ' + u + ' =====');
  fetch(u, { headers: { 'User-Agent': 'ds-harness' } })
    .then(function (r) { return r.text(); })
    .then(function (t) { console.log(t.slice(0, 5000)); })
    .catch(function (e) { console.log('ERR: ' + e.message); })
    .then(function () { go(i + 1); });
}
go(0);
