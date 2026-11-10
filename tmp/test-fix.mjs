import fs from 'fs';

const content = fs.readFileSync('/tmp/standalone-test.js', 'utf-8');
const target = "})()}\n                    ` : ''}";
console.log('Includes target:', content.includes(target));
const fixed = content.replace(target, "})()}");
fs.writeFileSync('/tmp/standalone-fixed.js', fixed);
console.log('Fixed written');
