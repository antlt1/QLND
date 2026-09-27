'use strict';
/** Doc Excel -> JSON -> Excel sach. Chay: qlnd.cmd build [--src <thu muc>] [--out <thu muc>] */
const { execFileSync } = require('child_process');
const path = require('path');

const args = process.argv.slice(2);
const run = (name) => {
  console.log(`\n=== ${name} ===`);
  execFileSync(process.execPath, [path.join(__dirname, name + '.js'), ...args], { stdio: 'inherit' });
};

run('extract');
run('excel');
console.log('\nXong.');
