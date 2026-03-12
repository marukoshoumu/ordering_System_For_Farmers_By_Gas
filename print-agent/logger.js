const fs = require('fs');
const path = require('path');

function createLogger(logDir) {
  if (!fs.existsSync(logDir)) {
    fs.mkdirSync(logDir, { recursive: true });
  }
  const getLogPath = () => {
    const date = new Date().toISOString().slice(0, 10);
    return path.join(logDir, `print-agent-${date}.log`);
  };
  const write = (level, message) => {
    const line = `${new Date().toISOString()}\t${level}\t${message}\n`;
    fs.appendFileSync(getLogPath(), line);
    console.log(line.trim());
  };
  return {
    info: (msg) => write('INFO', msg),
    warn: (msg) => write('WARN', msg),
    error: (msg) => write('ERROR', msg),
  };
}

module.exports = { createLogger };
