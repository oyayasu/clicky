import fs from 'node:fs';import path from 'node:path';
fs.rmSync('dist',{recursive:true,force:true});fs.mkdirSync('dist/src',{recursive:true});
fs.copyFileSync('index.html','dist/index.html');
for(const file of fs.readdirSync('src'))if(fs.statSync(path.join('src',file)).isFile())fs.copyFileSync(path.join('src',file),path.join('dist/src',file));
console.log('Clicky: arquivos estáticos preparados em dist/');
