// Run after npm ci: npm run verify:local-txt
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const path=require('node:path');
const ts=require('typescript');
const root=path.resolve(__dirname,'../app/queue');
const cache=new Map();
function load(name){
  if(cache.has(name))return cache.get(name).exports;
  const source=path.join(root,name+'.ts');
  const js=ts.transpileModule(fs.readFileSync(source,'utf8'),{fileName:source,compilerOptions:{target:ts.ScriptTarget.ES2020,module:ts.ModuleKind.CommonJS}}).outputText;
  const mod={exports:{}}; cache.set(name,mod);
  vm.runInNewContext('(function(require,module,exports){'+js+'})',{console}).call(null,relative=>load(relative.replace(/^\.\//,'')),mod,mod.exports);
  return mod.exports;
}
const {parseQueueTxt,formatQueueTxt}=load('local-text');
const kinds=parseQueueTxt('1. 그루밍 (유료)\n2. 프리모베일 (소개팅)\n3. 라미라미 (무료)\n4. Nuno_TV (멤버십)');
assert.equal(kinds.people.length,4);
assert.equal(JSON.stringify(kinds.people.map(p=>p.paid)), JSON.stringify([true,false,false,true]));
assert.equal(kinds.people[3].freeOrder,true);
assert.equal(JSON.stringify(parseQueueTxt(formatQueueTxt(kinds)).people.map(p=>p.name)), JSON.stringify(kinds.people.map(p=>p.name)));
const one=parseQueueTxt('\u200b[대기열] 1. 그루밍 (유료) - 2. 라미라미 (무료) (유료는 우선순위)');
assert.equal(one.people.length,2);
assert.equal(parseQueueTxt('').people.length,0);
assert.throws(()=>parseQueueTxt('1. A (무료)\n2. A (유료)'),/중복/);
console.log('Local TXT format tests passed');
