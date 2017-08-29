importScripts('xlsx.full.min.js');
postMessage({t:'ready'});
/*onmessage = function(evt) {
  var v;
  try { v = XLSX.read(evt.data.d, evt.data.b); }
  catch(e) { postMessage({t:"e",d:e.stack}); }
  postMessage({t:evt.data.t, d:JSON.stringify(v)});
}*/

onmessage = function (oEvent) {
    var v;
    try {
        v = XLSX.read(oEvent.data.d, {type: oEvent.data.b ? 'binary' : 'base64'});
    } catch(e) { postMessage({t:"e",d:e.stack||e}); }
    postMessage({t:"xlsx", d:JSON.stringify(v)});
};
