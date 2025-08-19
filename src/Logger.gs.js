// Logger.gs — tiny leveled logger for Apps Script

var LoggerEx = (function () {
  var LEVELS = { ERROR: 0, WARN: 1, INFO: 2, DEBUG: 3 };
  var current = LEVELS.INFO; // default: show INFO/WARN/ERROR

  function out(tag, level, args) {
    if (level > current) return;
    var msg = Array.prototype.map.call(args, function (x) {
      try {
        if (x instanceof Error) return x.stack || (x.name + ': ' + x.message);
        return typeof x === 'object' ? JSON.stringify(x) : String(x);
      } catch (_) { return String(x); }
    }).join(' ');
    Logger.log('[' + tag + '] ' + msg);
    // Uncomment if you also want console output (V8):
    // if (tag === 'ERROR') console.error(msg);
    // else if (tag === 'WARN') console.warn(msg);
    // else console.log(msg);
  }

  return {
    setLevel: function (levelName) { current = LEVELS[String(levelName).toUpperCase()] || current; },
    log:   function () { out('INFO',  LEVELS.INFO,  arguments); },
    warn:  function () { out('WARN',  LEVELS.WARN,  arguments); },
    error: function () { out('ERROR', LEVELS.ERROR, arguments); },
    debug: function () { out('DEBUG', LEVELS.DEBUG, arguments); }
  };
})();