const fs = require('fs');
const path = require('path');
const os = require('os');

console.log('Hello from Node.js!');
console.log(`Node version: ${process.version}`);
console.log(`Platform: ${process.platform} ${process.arch}`);

if (process.argv.includes('--diagnose')) {
  const outputDir = path.join(__dirname, 'output');
  fs.mkdirSync(outputDir, { recursive: true });
  const report = {
    app: 'Test_JS_Node',
    time: new Date().toISOString(),
    node: process.version,
    platform: process.platform,
    arch: process.arch,
    hostname: os.hostname(),
    cpus: os.cpus().length,
    totalMemoryGB: Math.round(os.totalmem() / 1024 / 1024 / 1024)
  };
  fs.writeFileSync(path.join(outputDir, 'node_diagnostic.json'), JSON.stringify(report, null, 2));
  console.log('Diagnostic written to output/node_diagnostic.json');
}
