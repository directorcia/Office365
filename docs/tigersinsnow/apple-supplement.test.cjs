'use strict';
const assert = require('node:assert/strict');
const test = require('node:test');
const game = require('./game.js');
const rules = game.manual;
const point = game.referenceMap.position;
const restore = state => {
  const copy = game.deserialize(JSON.stringify(state));
  assert.deepEqual(copy,state);
  return copy;
};
const seedForRoll = value => Number(((BigInt(value) - 1013904223n) * 4276115653n) & 0xffffffffn);

test('revised movement crosses rough river enemy zones and retains legacy save allowances', () => {
  for (let seed=1;seed<=20;seed++) {
    const state=game.createGame(seed,'manual','german',2,{playMode:'hotseat'});
    assert.equal(state.movementVersion,1);
    assert.ok(state.units.every(unit=>unit.mobility===(unit.type==='armor'?30:24)));
    assert.ok(state.units.filter(unit=>unit.side==='german').every(unit=>unit.mp===0));
    state.turn=2;state.weather='Overcast';
    rules.beginSide(state,'german');
    const unit=state.units.find(unit=>unit.name==='12th VG DIV');
    state.units.forEach(counter=>{counter.hp=0;});
    Object.assign(unit,point(16,4),{hp:unit.maxHp});
    const enemy=state.units.find(counter=>counter.side==='american');
    Object.assign(enemy,point(16,3),{hp:enemy.maxHp});
    rules.beginSide(state,'german');
    assert.equal(unit.mp,24);
    const destination=state.map[game.key(point(15,4).x,point(15,4).y)];
    const cost=rules.movementCost(state,unit,destination,unit);
    assert.ok(cost>=15&&cost<=19);
    unit.mp=8;
    const before=JSON.stringify(state);
    assert.match(rules.movementWarning(state,unit,destination),/8 enemy zone/);
    assert.equal(JSON.stringify(state),before);
    assert.equal(game.move(state,unit.id,destination.x,destination.y),false);
    unit.mp=24;
    assert.ok(game.reachable(state,unit).has(game.key(destination.x,destination.y)));
    assert.ok(game.move(state,unit.id,destination.x,destination.y));
    assert.equal(unit.mp,0);
    restore(state);
    const legacy=game.createGame(seed,'manual','german',2,{playMode:'hotseat'});
    delete legacy.movementVersion;
    legacy.units.forEach(counter=>{counter.mobility=counter.type==='armor'?10:8;});
    restore(legacy);
    const invalid=structuredClone(legacy);invalid.movementVersion=2;
    assert.throws(()=>game.deserialize(JSON.stringify(invalid)),/movement rules revision/);
  }
});

function advanceFixture(limit, roll, legacy=false) {
  const state = game.createGame(42,'manual','german',2,{appleSupplement:!legacy});
  state.units.forEach(unit => { unit.hp=0; });
  const attacker=state.units.find(unit=>unit.id==='card-g17'),defender=state.units.find(unit=>unit.id==='card-a2');
  Object.assign(attacker,point(5,0),{hp:115,fired:true});
  if(!legacy) attacker.combatCount=1;
  Object.assign(defender,point(6,0));
  state.seed=seedForRoll(roll);
  const battle={ids:[attacker.id],defenderId:defender.id,
    before:[{id:attacker.id,name:attacker.name,side:attacker.side,hp:115},
      {id:defender.id,name:defender.name,side:defender.side,hp:75}],
    original:{x:defender.x,y:defender.y},attackerPositions:[{id:attacker.id,x:attacker.x,y:attacker.y}],
    odds:3,attackStrategy:'assault',defenseStrategy:'hold',artillery:0,defensiveArtillery:0,
    events:[],retreatCount:0,advanceLimit:limit,attackerMustRetreat:false};
  return {state,attacker,defender,battle};
}

test('supplement second combat is certain for three hexes, exactly 25% for two, and announced before advance', () => {
  for(const [limit,roll,eligible] of [[1,0,false],[2,1073741823,true],[2,1073741824,false],[3,4294967295,true]]) {
    let {state,attacker,battle}=advanceFixture(limit,roll);
    const result=rules.maneuvers.start(state,battle);
    assert.equal(result.pending,true);
    assert.deepEqual(state.pendingBattle.secondCombatIds,eligible?[attacker.id]:[]);
    assert.equal(attacker.fired,true);
    assert.equal(state.pendingBattle.events.some(event=>event.includes('Units eligible for second combat')),eligible);
    const seed=state.seed;
    state=restore(state);
    const copy=restore(state);
    rules.chooseManeuver(state,'finish');rules.chooseManeuver(copy,'finish');
    assert.deepEqual(copy,state);
    assert.equal(state.units.find(unit=>unit.id===attacker.id).fired,!eligible);
    assert.equal(state.seed,seed);
  }
  let eligible=0;
  for(let i=0;i<100;i++) {
    const {state,battle}=advanceFixture(2,Math.floor(i/100*4294967296));
    rules.maneuvers.start(state,battle);
    if(state.pendingBattle.secondCombatIds.length) eligible++;
  }
  assert.equal(eligible,25);
});

test('second-combat eligibility rewinds computer order stepping and rejects corrupt pending saves', () => {
  let {state,battle,attacker}=advanceFixture(3,0);
  state.computerProgress={stage:'combat',index:17};
  rules.maneuvers.start(state,battle);
  const invalid=structuredClone(state);
  invalid.pendingBattle.secondCombatIds=['card-a2'];
  assert.throws(()=>game.deserialize(JSON.stringify(invalid)),/eligibility/);
  delete state.computerProgress;
  state=restore(state);
  state.computerProgress={stage:'combat',index:17};
  rules.chooseManeuver(state,'finish');
  assert.equal(state.computerProgress.index,16);
  assert.equal(state.units.find(unit=>unit.id===attacker.id).fired,false);
  const fixture=advanceFixture(3,0);
  fixture.attacker.combatCount=2;
  rules.maneuvers.start(fixture.state,fixture.battle);
  assert.deepEqual(fixture.state.pendingBattle.secondCombatIds,[]);
  rules.chooseManeuver(fixture.state,'finish');
  assert.equal(fixture.attacker.fired,true);
});

test('elimination risk includes 25 CP but excludes 26 CP and zero-loss events', () => {
  for(const [hp,loss,expected] of [[26,1,0],[27,1,26],[25,0,25]]) {
    const state=game.createGame(42,'manual');
    const unit=state.units[0];unit.hp=hp;state.seed=seedForRoll(0);
    rules.lose(state,unit,loss);
    assert.equal(unit.hp,expected);
    assert.equal(state.casualties.german,hp-expected);
    assert.equal(state.victoryPoints.american,hp-expected);
  }
  const state=game.createGame(42,'manual','german',2,{appleSupplement:false});
  state.units[0].hp=26;state.seed=seedForRoll(0);
  rules.lose(state,state.units[0],1);
  assert.equal(state.units[0].hp,25);
  restore(state);
});

test('weighted attack pressure uses major:medium:light:recon 9:3:1:1 and resets per faction turn', () => {
  const state=game.createGame(42,'manual');
  assert.equal(rules.alteredStrategyChance(state,'german','major'),0.06);
  state.strategyPressure.german=18;
  assert.ok(Math.abs(rules.alteredStrategyChance(state,'german','major')-0.16)<1e-12);
  const medium=rules.alteredStrategyChance(state,'german','assault')-0.04;
  const light=rules.alteredStrategyChance(state,'german','probe')-0.024;
  const recon=rules.alteredStrategyChance(state,'german','recon')-0.014;
  assert.ok(Math.abs(medium-3*light)<1e-12);
  assert.ok(Math.abs(light-recon)<1e-12);
  assert.ok(rules.alteredStrategyChance(state,'german','major')>0.06);
  const before=JSON.stringify(state);
  assert.throws(()=>rules.resolveAttack(state,['unknown'],'card-a2',{strategy:'major'}),/eligible/);
  assert.equal(JSON.stringify(state),before);
  state.strategyPressure.american=9;
  rules.beginSide(state,'american');
  assert.equal(state.strategyPressure.american,0);
  assert.equal(state.strategyPressure.german,18);
  state.playMode='hotseat';state.playerSide='american';
  restore(state);
  for(const [strategy,weight] of [['recon',1],['probe',1],['assault',3],['major',9]]) {
    const fixture=advanceFixture(1,2147483648);
    fixture.attacker.fired=false;fixture.attacker.combatCount=0;fixture.defender.hp=75;
    const before=JSON.stringify(fixture.state);
    rules.preview(fixture.state,[fixture.attacker.id],fixture.defender.id,{strategy});
    assert.equal(JSON.stringify(fixture.state),before);
    rules.resolveAttack(fixture.state,[fixture.attacker.id],fixture.defender.id,{strategy});
    assert.equal(fixture.state.strategyPressure.german,weight);
    assert.equal(fixture.attacker.combatCount,1);
  }
});

test('air interdiction reduces actual fuel, supply and MP once per day and reports formation fuel shortage', () => {
  const state=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
  state.turn=8;state.weather='Clear';state.fuel=0;
  const supply=state.generalSupply.german;
  rules.beginSide(state,'german');
  assert.equal(state.fuel,0);
  assert.equal(state.generalSupply.german,supply-state.airInterdiction.supply);
  assert.equal(state.airInterdiction.turn,8);
  assert.ok(Object.values(state.airInterdiction.movement).some(loss=>loss>0));
  assert.ok(state.reports.german.some(event=>event.text.includes('unit out of fuel')));
  const air=structuredClone(state.airInterdiction),remaining=state.generalSupply.german;
  rules.beginSide(state,'german');
  assert.equal(state.generalSupply.german,remaining);
  assert.equal(state.airInterdiction.fuel,air.fuel);
  assert.equal(state.airInterdiction.supply,air.supply);
  const report=rules.phaseReport(state);
  assert.equal(report.armies[1].fuel,state.fuel);
  assert.equal(report.armies[1].supply,state.generalSupply.german);
  assert.equal(report.armies[0].fuel,null);
  assert.equal(report.score,state.score);
  report.interdiction.supply=999;
  assert.notEqual(state.airInterdiction.supply,999);
  restore(state);
});

test('Allied bridge warning identifies an unaffordable adjacent river crossing without spending resources', () => {
  const state=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
  game.endTurn(state,'german');
  state.units.forEach(unit=>{unit.hp=0;});
  const unit=state.units.find(unit=>unit.id==='card-a1');
  unit.hp=unit.maxHp;unit.mp=1;
  let destination;
  for(const tile of Object.values(state.map)) {
    destination=game.neighbors(state,tile).find(next=>game.referenceMap.crossesRiver(tile,next));
    if(destination) {unit.x=tile.x;unit.y=tile.y;break;}
  }
  assert.ok(destination);
  const before=structuredClone(state);
  assert.match(rules.movementWarning(state,unit,destination),/traffic jam on bridge/);
  assert.equal(game.move(state,unit.id,destination.x,destination.y),false);
  assert.match(state.log[0],/traffic jam on bridge/);
  before.log=structuredClone(state.log);before.reports=structuredClone(state.reports);
  assert.deepEqual(state,before);
  unit.mp=24;
  assert.equal(rules.movementWarning(state,unit,destination),'');
});

test('supplement saves retain pressure and interdiction; older card and Classic saves stay unchanged', () => {
  for(const options of [{},{appleSupplement:false}]) {
    const state=game.createGame(42,'manual','german',2,options);
    assert.equal(rules.supplement(state),options.appleSupplement!==false);
    restore(state);
  }
  const state=game.createGame(42,'manual');
  for(const mutate of [
    s=>{s.appleSupplementVersion=2;},s=>{s.strategyPressure.german=-1;},
    s=>{s.strategyPressure.german=0.5;},s=>{delete s.airInterdiction;},
    s=>{s.airInterdiction.fuel=6;},s=>{s.airInterdiction.movement.unknown=1;},
    s=>{s.units[0].combatCount=3;},
    s=>{delete s.appleSupplementVersion;}
  ]) {
    const invalid=structuredClone(state);mutate(invalid);
    assert.throws(()=>game.deserialize(JSON.stringify(invalid)),/supplement|interdiction|unit data/i);
  }
  const classic=game.createGame(42);
  assert.equal(classic.appleSupplementVersion,undefined);
  restore(classic);
});
