const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const root=path.resolve(__dirname,'..');
const ts=require(path.join(root,'node_modules/typescript'));
const source=fs.readFileSync(path.join(root,'app/page.tsx'),'utf8');
const ast=ts.createSourceFile('page.tsx',source,ts.ScriptTarget.Latest,true,ts.ScriptKind.TSX);
const declarations=new Map();
function visit(node){if(ts.isVariableDeclaration(node)&&ts.isIdentifier(node.name)&&node.initializer)declarations.set(node.name.text,node.initializer.getText(ast));ts.forEachChild(node,visit);}visit(ast);
const names=['DATING_TEN_GOD_GROUPS','BRANCH_RELATION_RULES','STEM_HAP_PAIRS','STEM_CHUNG_PAIRS','getDatingAnyPillarBranch','getDatingPillarBranch','getDatingDayStem','getDatingPillarGanji','isDatingPillarDayMatch','isDatingBranchRelation','isDatingStemHap','getCheoneulTargetBranches','getAllDatingBranches','isDatingCheoneulMatch','getDatingTenGodCounts','getDatingElementCount','parseDatingFilterCount','normalizeDatingStars','getDatingMatchMeta','restoreDatingCountRules','normalizeDatingTenGodPreset','matchesDatingRelationFilter','filteredDatingCandidates'];
const code=ts.transpileModule(names.map(name=>`const ${name}=${declarations.get(name)};`).join('\n'),{compilerOptions:{target:ts.ScriptTarget.ES2020}}).outputText;
const branches=['자','축','인','묘','진','사','오','미','신','유','술','해'];
const bH=['子','丑','寅','卯','辰','巳','午','未','申','酉','戌','亥'];
const stems=['갑','을','병','정','무','기','경','신','임','계'];
const sH=['甲','乙','丙','丁','戊','己','庚','辛','壬','癸'];
const norm=(kor,hanja)=>value=>kor[hanja.indexOf(String(value||''))]||String(value||'');
const normalizeBranch=norm(branches,bH),normalizeStem=norm(stems,sH);
const pair=(a,b)=>[normalizeBranch(a),normalizeBranch(b)].sort().join('-');
const defaults={dayRelations:[],dayRelationJoin:'or',excludeDayWonjin:false,monthRelations:[],monthRelationJoin:'or',excludeMonthWonjin:false,stemHapMode:'off',dayStems:[],cheoneulMode:'off',pillarDayMatch:false,excludeSameDayPillar:false,genderMode:'all',olderMaxAgeGap:'',youngerMaxAgeGap:'',minStars:0,tenGodMinRules:[],tenGodMaxRules:[],excludeMissingTenGods:[],tenGodMajorities:[],elementMinRules:[],elementMaxRules:[],sortMode:'default'};
const person=(id,year=1995)=>({id,name:id,birthDate:`${year}-06-15`,calendarType:'solar',gender:'여성'});
const pillar=(stem,branch)=>({stem:sH[stem],branch:bH[branch],stemIndex:stem,branchIndex:branch,ganji:sH[stem]+bH[branch]});
function chart(day=0,month=2,year=1995){return{year:pillar(0,8),month:pillar(2,month),day:pillar(0,day),hour:null,solarBirthDate:`${year}-06-15`,tenGods:{yearStem:'비견',yearBranch:'편인',monthStem:'식신',monthBranch:'비견',dayStem:'일간',dayBranch:'정인',hourStem:'',hourBranch:''},elementCount:{wood:3,fire:1,earth:0,metal:1,water:1}};}
function harness({filters={},candidates=[person('candidate')],fixed=person('fixed'),fixedChart=chart(),charts,meta={}}={}){
  const cache=charts||new Map(candidates.map(p=>[p.id,chart()]));
  const env={BRANCHES:branches,normalizeBranch,normalizeStem,getItemBranch:item=>normalizeBranch(item.data?.branch),makeBranchPairKey:pair,makeHanjaPairKey:(a,b)=>[a,b].sort().join('-'),datingFixedPerson:fixed,datingMatchMeta:meta,makeDatingMatchKey:(a,b)=>a.id+'>>'+b.id,makePersonKey:p=>p.id,datingFilterRuleIdRef:{current:1},DATING_TEN_GOD_FILTER_OPTIONS:['비겁','식상','재성','관성','인성'],DATING_ELEMENT_FILTER_OPTIONS:['목','화','토','금','수'],datingAvailableCandidates:candidates,datingFixedSajuForFilter:fixedChart,datingSajuCache:cache,datingFilterSnapshot:{...defaults,...filters},useMemo:fn=>fn(),normalizeDatingPerson:p=>p,calculateOneSaju:()=>{throw new Error('failed cached calculation retried during filtering');}};
  return new Function(...Object.keys(env),code+';return {result:filteredDatingCandidates,getDatingMatchMeta,restoreDatingCountRules,normalizeDatingTenGodPreset,matchesDatingRelationFilter,getDatingTenGodCounts};')(...Object.values(env));
}
let failed=0,passed=0;function test(name,fn){try{fn();passed++;}catch(e){failed++;console.log('FAIL:',name,'-',e.message);}}
const result=options=>harness(options).result.map(p=>p.id);
test('OFF preserves candidate order',()=>assert.deepEqual(result(),['candidate']));
test('one failed cached candidate cannot crash all candidates',()=>assert.deepEqual(result({filters:{dayStems:['갑']},charts:new Map([['candidate',null]])}),[]));
test('corrupt stars cannot pass minimum',()=>assert.deepEqual(result({filters:{minStars:1},meta:{'fixed>>candidate':{stars:'bad'}}}),[]));
test('unknown-time chart counts six symbols',()=>assert.equal(Object.values(harness().getDatingTenGodCounts(chart())).slice(0,10).reduce((a,b)=>a+b,0),6));
test('Samhap only compares the selected pillars', () => {
  const fixedChart = { ...chart(11, 1), year: pillar(6, 4), hour: pillar(2, 6) };
  const candidate = { ...chart(1, 7), year: pillar(2, 2), hour: pillar(5, 3) };
  const filters = { dayRelations: ['삼합'], monthRelations: ['삼합'], excludeDayWonjin: true, excludeMonthWonjin: true };
  const matches = (chart, activeFilters = filters) => result({ fixedChart, filters: activeFilters, charts: new Map([['candidate', chart]]) });
  assert.deepEqual(matches(candidate), []);
  assert.deepEqual(matches({ ...candidate, day: pillar(1, 3) }, { dayRelations: ['삼합'] }), ['candidate']);
  assert.deepEqual(matches({ ...candidate, month: pillar(1, 9) }, { monthRelations: ['삼합'] }), ['candidate']);
});
test('Samhap matches both same-position pairs without the central branch', () => {
  const fixedChart = { ...chart(5, 7), year: pillar(9, 9), hour: pillar(7, 11) };
  const candidate = { ...chart(1, 11), year: pillar(1, 11), hour: pillar(2, 0) };
  const filters = { dayRelations: ['삼합'], monthRelations: ['삼합'], dayMonthJoin: 'and' };
  const matches = candidateChart => result({ fixedChart, filters, charts: new Map([['candidate', candidateChart]]) });
  assert.deepEqual(matches(candidate), ['candidate']);
  assert.deepEqual(matches({ ...candidate, month: pillar(0, 0) }), []);
  assert.deepEqual(matches({ ...candidate, day: fixedChart.day }), []);
});
test('day/month blocks use OR, exclusions stay mandatory',()=>{const c=chart(1,8);assert.deepEqual(result({filters:{dayRelations:['육합'],monthRelations:['육합']},charts:new Map([['candidate',c]])}),['candidate']);c.month=pillar(0,11);const fixed=chart(0,8);assert.deepEqual(result({fixedChart:fixed,filters:{dayRelations:['육합'],excludeMonthWonjin:true},charts:new Map([['candidate',{...c,month:pillar(0,3)}]])}),[]);});
test('day/month AND requires both active cards to match', () => {
  const filters = { dayRelations: ['육합'], monthRelations: ['육합'], dayMonthJoin: 'and' };
  const matches = (candidate, activeFilters = filters) => result({ filters: activeFilters, charts: new Map([['candidate', candidate]]) });
  assert.deepEqual(matches(chart(1, 8)), []);
  assert.deepEqual(matches(chart(1, 11)), ['candidate']);
  assert.deepEqual(matches(chart(1, 8), { ...filters, monthRelations: [] }), ['candidate']);
});
test('per-card OR and AND differ',()=>{const h=harness();assert.equal(h.matchesDatingRelationFilter('자','축',['육합','삼합'],'or'),true);assert.equal(h.matchesDatingRelationFilter('자','축',['육합','삼합'],'and'),false);});
test('same-day exclusion overrides pillar match',()=>assert.deepEqual(result({filters:{pillarDayMatch:true,excludeSameDayPillar:true}}),[]));
test('count min/max remain AND',()=>assert.deepEqual(result({filters:{tenGodMinRules:[{target:'비겁',count:'4'}],elementMaxRules:[{target:'목',count:'2'}]}}),[]));
test('invalid preset targets cannot exclude everyone',()=>assert.deepEqual(harness().restoreDatingCountRules([{target:'removed-option',count:'2'}],['목','화','토','금','수']),[]));
test('lunar age gap uses converted solar year',()=>assert.deepEqual(result({fixed:{...person('fixed',1994),calendarType:'lunar'},fixedChart:chart(0,2,1995),candidates:[person('candidate',1995)],filters:{youngerMaxAgeGap:'0'}}),['candidate']));
test('invalid birth year cannot pass an active age restriction',()=>assert.deepEqual(result({candidates:[{...person('candidate'),birthDate:''}],charts:new Map([['candidate',null]]),filters:{youngerMaxAgeGap:'0'}}),[]));
test('score sorting places blank after numeric zero',()=>assert.deepEqual(result({candidates:[person('blank'),person('zero')],filters:{sortMode:'score-desc'},meta:{'fixed>>blank':{score:' '},'fixed>>zero':{score:'0'}}}),['zero','blank']));
test('stable score sorting preserves ties',()=>assert.deepEqual(result({candidates:[person('b'),person('a')],filters:{sortMode:'score-desc'},meta:{'fixed>>b':{score:'90'},'fixed>>a':{score:'90'}}}),['b','a']));
test('opposite-gender filter excludes same gender',()=>assert.deepEqual(result({filters:{genderMode:'exclude-same'}}),[]));
test('same-gender filter accepts same gender',()=>assert.deepEqual(result({filters:{genderMode:'same-only'}}),['candidate']));
test('opposite-gender filter accepts different gender',()=>assert.deepEqual(result({candidates:[{...person('candidate'),gender:'남성'}],filters:{genderMode:'exclude-same'}}),['candidate']));
test('older-age maximum includes its boundary',()=>assert.deepEqual(result({candidates:[person('candidate',1992)],charts:new Map([['candidate',chart(0,2,1992)]]),filters:{olderMaxAgeGap:'3'}}),['candidate']));
test('older-age maximum rejects above boundary',()=>assert.deepEqual(result({candidates:[person('candidate',1991)],charts:new Map([['candidate',chart(0,2,1991)]]),filters:{olderMaxAgeGap:'3'}}),[]));
test('zero older limit leaves unrestricted younger direction',()=>assert.deepEqual(result({candidates:[person('candidate',1996)],charts:new Map([['candidate',chart(0,2,1996)]]),filters:{olderMaxAgeGap:'0'}}),['candidate']));
test('empty maximum is inactive',()=>assert.deepEqual(result({filters:{elementMaxRules:[{target:'목',count:''}]}}),['candidate']));
test('zero maximum really excludes a nonzero element',()=>assert.deepEqual(result({filters:{elementMaxRules:[{target:'목',count:'0'}]}}),[]));
test('unknown-hour branches cannot fabricate Cheoneul match',()=>assert.deepEqual(result({fixedChart:chart(),filters:{cheoneulMode:'include'}}),[]));
test('finite star normalization and clamping',()=>{for(const [value,expected] of [['bad',0],[Infinity,0],[-1,0],[9,3],[2.8,2]])assert.equal(harness({meta:{'fixed>>candidate':{stars:value}}}).getDatingMatchMeta(person('candidate')).stars,expected);});
test('all 144 seasonal-pair combinations match distinct 2-symbol rule',()=>{
 const h=harness(); const seasons=[[2,3,4],[5,6,7],[8,9,10],[11,0,1]];
 for(let a=0;a<12;a++)for(let b=0;b<12;b++) assert.equal(h.matchesDatingRelationFilter(branches[a],branches[b],['방합'],'or'),a!==b&&seasons.some(group=>group.includes(a)&&group.includes(b)));
});
test('Banghap compares only the selected matching pillar', () => {
  const fixedChart = { ...chart(10, 5), year: pillar(6, 4), hour: pillar(7, 7) };
  const candidate = { ...chart(7, 0), year: pillar(3, 1), hour: pillar(0, 4) };
  const matches = (candidateChart, filters) => result({ fixedChart, filters, charts: new Map([['candidate', candidateChart]]) });
  assert.deepEqual(matches(candidate, { dayRelations: ['방합'], monthRelations: ['방합'] }), []);
  assert.deepEqual(matches({ ...candidate, day: pillar(0, 8) }, { dayRelations: ['방합'] }), ['candidate']);
  assert.deepEqual(matches({ ...candidate, month: pillar(0, 7) }, { monthRelations: ['방합'] }), ['candidate']);
});
test('zero Ten-God maximum accepts absence and rejects one occurrence',()=>{
 const filters={tenGodMaxRules:[{target:'관성',count:'0'}]};
 assert.deepEqual(result({filters}),['candidate']);
 const present=chart();present.tenGods.yearStem='편관';
 assert.deepEqual(result({filters,charts:new Map([['candidate',present]])}),[]);
});
test('zero grouped Ten-God maximum tests both members',()=>{
 const filters={tenGodMaxRules:[{target:'재성',count:'0'}]};
 assert.deepEqual(result({filters}),['candidate']);
 for(const member of ['정재','편재']){const present=chart();present.tenGods.yearStem=member;assert.deepEqual(result({filters,charts:new Map([['candidate',present]])}),[]);}
});
test('zero element maximum accepts absence and rejects one occurrence',()=>{
 const filters={elementMaxRules:[{target:'토',count:'0'}]};
 assert.deepEqual(result({filters}),['candidate']);
 const present=chart();present.elementCount.earth=1;
 assert.deepEqual(result({filters,charts:new Map([['candidate',present]])}),[]);
});
test('multiple zero maxima require every requested component absent',()=>{
 const filters={tenGodMaxRules:[{target:'관성',count:'0'},{target:'재성',count:'0'}],elementMaxRules:[{target:'토',count:'0'}]};
 assert.deepEqual(result({filters}),['candidate']);
 const present=chart();present.tenGods.yearStem='정재';
 assert.deepEqual(result({filters,charts:new Map([['candidate',present]])}),[]);
});
test('blank maximum differs from explicit zero',()=>{
 assert.deepEqual(result({filters:{tenGodMaxRules:[{target:'비겁',count:''}]}}),['candidate']);
 assert.deepEqual(result({filters:{tenGodMaxRules:[{target:'비겁',count:'0'}]}}),[]);
});
test('legacy Ten-God preset targets migrate to their groups', () => {
  const normalized = harness().normalizeDatingTenGodPreset({
    tenGodMinRules: [{ target: '편재', count: 2 }, { target: '정관', count: '1' }],
    tenGodMaxRules: [{ target: '정재', count: '0' }, { target: '식신', count: '0' }],
  });
  assert.deepEqual(normalized.tenGodMinRules, [{ target: '재성', count: '2' }, { target: '관성', count: '1' }]);
  assert.deepEqual(normalized.tenGodMaxRules, [{ target: '재성', count: '0' }, { target: '식상', count: '0' }]);
});
test('missing Ten-God filter requires zero of every selected group', () => {
  const filters = { excludeMissingTenGods: ['비겁', '식상'] };
  assert.deepEqual(result({ filters }), []);
  const candidate = chart();
  candidate.tenGods.yearStem = '편재';
  candidate.tenGods.monthStem = '편재';
  const matches = () => result({ filters, charts: new Map([['candidate', candidate]]) });
  assert.deepEqual(matches(), []);
  candidate.tenGods.monthBranch = '정재';
  assert.deepEqual(matches(), ['candidate']);
});
test('Ten-God majority requires at least three of each selected group', () => {
  const candidate = chart();
  candidate.tenGods.hourStem = '식신';
  const matches = (tenGodMajorities) => result({ filters: { tenGodMajorities }, charts: new Map([['candidate', candidate]]) });
  assert.deepEqual(matches(['식상']), []);
  candidate.tenGods.hourBranch = '상관';
  assert.deepEqual(matches(['식상', '비겁']), ['candidate']);
  assert.deepEqual(matches(['식상', '재성']), []);
});
test('Ten-God majority recognizes a matching month pillar and doubles the day branch', () => {
  const matches = candidate => result({ filters: { tenGodMajorities: ['식상'] }, charts: new Map([['candidate', candidate]]) });
  const monthPillar = chart();
  monthPillar.tenGods.monthBranch = '식신';
  assert.deepEqual(matches(monthPillar), ['candidate']);
  const dayBranch = chart();
  dayBranch.tenGods.dayBranch = '상관';
  assert.deepEqual(matches(dayBranch), ['candidate']);
  dayBranch.tenGods.monthStem = '편인';
  assert.deepEqual(matches(dayBranch), []);
  assert.deepEqual(result({ filters: { tenGodMinRules: [{ target: '식상', count: '3' }] }, charts: new Map([['candidate', monthPillar]]) }), []);
});
console.log(`${passed} passed, ${failed} failed`);process.exitCode=failed?1:0;
