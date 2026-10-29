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

test('Meuse edges match the supplied red course without including other river branches', () => {
  const reference=game.referenceMap;
  assert.equal(reference.meuseEdges.length,31);
  assert.equal(new Set(reference.meuseEdges.map(({a,b})=>reference.edgeKey(a,b))).size,31);
  for(const {a,b} of reference.meuseEdges) {
    assert.equal(game.distance(a,b),1);
    assert.equal(reference.crossesRiver(a,b),true);
    assert.equal(reference.crossesMeuse(a,b),true);
    assert.equal(reference.crossesMeuse(b,a),true);
  }
  for(const [ac,ar,bc,br] of [[0,2,1,3],[8,2,9,2],[10,2,11,2],[3,5,3,6],[13,6,14,6],[1,8,2,8]]) {
    const a=point(ac,ar),b=point(bc,br);
    assert.equal(reference.crossesRiver(a,b),true);
    assert.equal(reference.crossesMeuse(a,b),false);
  }
});

function britishFixture() {
  const state=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
  state.turn=9;state.units.forEach(unit=>{unit.hp=0;});
  rules.reinforce(state);
  state.units.forEach(unit=>{unit.hp=0;});
  state.phase='american';state.playerSide='american';state.orderPhase='movement';
  return state;
}

test('rule 17.1 blocks all six British formations at the Meuse but permits other crossings', () => {
  const state=britishFixture();
  const british=state.units.filter(unit=>unit.nationality==='british');
  assert.equal(british.length,6);
  const us=state.units.find(unit=>unit.nationality==='us');
  for(const unit of [...british,us]) {
    for(const [ac,ar,bc,br,meuse] of [[1,4,2,4,true],[2,4,1,4,true],[8,2,9,2,false],[9,2,8,2,false],[1,4,1,5,false]]) {
      state.units.forEach(unit=>{unit.hp=0;});
      Object.assign(unit,point(ac,ar),{hp:unit.maxHp});
      const destination=state.map[game.key(point(bc,br).x,br)];
      unit.mp=rules.movementCost(state,unit,destination,unit);
      const allowed=unit.nationality!=='british'||!meuse;
      const before={...unit},fuel=state.fuel,supply={...state.generalSupply};
      assert.equal(game.reachable(state,unit).has(game.key(destination.x,destination.y)),allowed);
      assert.equal(game.move(state,unit.id,destination.x,destination.y),allowed);
      if(!allowed) {
        assert.deepEqual(unit,before);
        assert.equal(state.fuel,fuel);
        assert.deepEqual(state.generalSupply,supply);
        assert.match(state.log[0],/British.*Meuse.*17\.1/);
      }
    }
  }
});

test('British route planning never crosses any tagged Meuse edge', () => {
  const state=britishFixture(),reference=game.referenceMap;
  const unit=state.units.find(unit=>unit.nationality==='british'&&unit.type==='armor');
  unit.hp=unit.maxHp;unit.mp=32;
  for(const {a,b} of reference.meuseEdges) {
    for(const from of [a,b]) {
      Object.assign(unit,from);
      for(const option of game.reachable(state,unit).values()) {
        let previous=unit;
        for(const next of option.path) {
          assert.equal(reference.crossesMeuse(previous,next),false);
          previous=next;
        }
      }
    }
  }
});

test('British retreats and advances cannot cross the Meuse but can cross other rivers', () => {
  for(const kind of ['retreat','advance']) {
    for(const [ac,ar,bc,br,meuse] of [[1,4,2,4,true],[2,4,1,4,true],[8,2,9,2,false],[9,2,8,2,false]]) {
      const state=britishFixture();state.orderPhase='combat';
      const unit=state.units.find(unit=>unit.nationality==='british');
      const enemy=state.units.find(unit=>unit.side==='german');
      Object.assign(unit,point(ac,ar),{hp:unit.maxHp});
      Object.assign(enemy,kind==='retreat'?point(15,0):point(bc,br),{hp:kind==='retreat'?enemy.maxHp:0});
      const destination=state.map[game.key(point(bc,br).x,br)];
      const attacker=kind==='retreat'?enemy:unit,defender=kind==='retreat'?unit:enemy;
      state.phase=attacker.side;state.playerSide=attacker.side;
      let result=rules.maneuvers.start(state,{ids:[attacker.id],defenderId:defender.id,
        before:[attacker,defender].map(counter=>({id:counter.id,name:counter.name,side:counter.side,hp:counter.maxHp})),
        original:{x:defender.x,y:defender.y},attackerPositions:[{id:attacker.id,x:attacker.x,y:attacker.y}],
        odds:2,attackStrategy:'assault',defenseStrategy:'hold',artillery:0,defensiveArtillery:0,
        events:[],retreatCount:kind==='retreat'?1:0,advanceLimit:kind==='advance'?1:0,attackerMustRetreat:false});
      if(kind==='advance'&&meuse) {
        assert.equal(state.pendingBattle,undefined);
        assert.deepEqual(result.advancePaths,[]);
        assert.equal(game.key(unit.x,unit.y),game.key(point(ac,ar).x,ar));
        continue;
      }
      const choices=rules.maneuvers.choices(state);
      const crossing=choices.find(choice=>game.key(choice.tile.x,choice.tile.y)===game.key(destination.x,destination.y));
      assert.equal(Boolean(crossing),!meuse);
      if(meuse) {
        const direction=game.aidCard.directions.findIndex(([dx,dy])=>unit.x+dx===destination.x&&unit.y+dy===destination.y)+1;
        const before=JSON.stringify(state);
        assert.throws(()=>rules.chooseManeuver(state,direction),/blocked/);
        assert.equal(JSON.stringify(state),before);
      }
      const choice=crossing||choices[0];
      result=rules.chooseManeuver(state,choice.direction);
      const path=kind==='retreat'?result.retreatPath:result.advancePaths[0].path;
      const next=state.map[path[0]];
      assert.equal(game.referenceMap.crossesMeuse(point(ac,ar),next),false);
    }
  }
});

test('new supplement resources use example-scale stocks and survive recovery and saves', () => {
  const state=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
  assert.equal(state.resourceVersion,1);
  assert.deepEqual(state.generalSupply,{german:1577,american:828});
  assert.equal(state.fuel,1023);
  assert.deepEqual(state.artillery,{german:8,american:5});
  restore(state);
  state.artillery={german:0,american:0};
  rules.finishDay(state);
  assert.equal(state.turn,2);
  assert.deepEqual(state.generalSupply,{german:1647,american:864});
  assert.deepEqual(state.artillery,{german:8,american:5});
  assert.equal(state.fuel,1023);
  state.defensiveArtillery=8;
  restore(state);
  for(const mutate of [s=>{s.resourceVersion=2;},s=>{s.generalSupply.german=-1;},
    s=>{s.generalSupply.american=1.5;},s=>{s.fuel=1024;},s=>{s.artillery.german=9;}]) {
    const invalid=structuredClone(state);mutate(invalid);
    assert.throws(()=>game.deserialize(JSON.stringify(invalid)),/Invalid|revision/);
  }
});

test('older supplement saves retain the original resource scale and daily caps', () => {
  let state=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
  delete state.resourceVersion;
  state.generalSupply={german:120,american:120};state.fuel=120;
  state.artillery={german:6,american:4};
  state=restore(state);
  rules.finishDay(state);
  assert.deepEqual(state.generalSupply,{german:120,american:120});
  assert.deepEqual(state.artillery,{german:6,american:4});
  assert.equal(state.fuel,120);
  restore(state);
});

test('rule 5.1 movement rolls and Allied starting-zone deductions are separate from entry costs', () => {
  for(const side of ['german','american']) {
    for(const type of ['infantry','armor']) {
      for(const adjacent of [false,true]) {
        for(let roll=0;roll<5;roll++) {
          const state=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
          state.turn=2;state.weather='Overcast';
          state.units.forEach(unit=>{unit.hp=0;});
          const unit=state.units.find(unit=>unit.side===side&&unit.type===type);
          Object.assign(unit,point(6,4),{hp:unit.maxHp});
          if(adjacent) {
            const enemy=state.units.find(unit=>unit.side!==side);
            Object.assign(enemy,point(7,4),{hp:enemy.maxHp});
          }
          const target=Math.floor((roll+0.5)/5*4294967296);
          state.seed=seedForRoll(target);
          if(!adjacent) state.seed=seedForRoll(state.seed);
          rules.beginSide(state,side);
          assert.equal(unit.supplyState,'supplied');
          assert.equal(unit.mp,unit.mobility+roll-2-(side==='american'&&adjacent?4:0));
          const destination=state.map[game.key(point(7,5).x,5)];
          const cost=rules.movementCost(state,unit,destination,unit);
          const clearState={...state,units:state.units.filter(unit=>unit.side===side)};
          assert.equal(cost-rules.movementCost(clearState,unit,destination,unit),adjacent?8:0);
          restore(state);
        }
      }
    }
  }
});

test('movement revision 1 retains fixed starting allowances beside enemies', () => {
  const state=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
  state.movementVersion=1;state.turn=2;state.weather='Overcast';
  state.units.forEach(unit=>{unit.hp=0;});
  const unit=state.units.find(unit=>unit.side==='american'&&unit.type==='infantry');
  const enemy=state.units.find(unit=>unit.side==='german');
  Object.assign(unit,point(6,4),{hp:unit.maxHp});
  Object.assign(enemy,point(7,4),{hp:enemy.maxHp});
  rules.beginSide(state,'american');
  assert.equal(unit.mp,24);
  restore(state);
});

test('revised movement crosses rough river enemy zones and retains legacy save allowances', () => {
  for (let seed=1;seed<=20;seed++) {
    const state=game.createGame(seed,'manual','german',2,{playMode:'hotseat'});
    assert.equal(state.movementVersion,2);
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
    assert.ok(unit.mp>=22&&unit.mp<=26);
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
    state.seed=seedForRoll(Math.floor(0.9*4294967296));
    assert.ok(game.move(state,unit.id,destination.x,destination.y));
    assert.equal(unit.mp,0);
    restore(state);
    const legacy=game.createGame(seed,'manual','german',2,{playMode:'hotseat'});
    delete legacy.movementVersion;
    legacy.units.forEach(counter=>{counter.mobility=counter.type==='armor'?10:8;});
    restore(legacy);
    const invalid=structuredClone(legacy);invalid.movementVersion=3;
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

test('rule 7.8 worked example: 11:1 Major versus Delaying costs 0/13 CP and retreats three hexes', () => {
  const state=game.createGame(42,'manual','german',2,{playMode:'hotseat',ratings:{german:7,american:5}});
  state.units.forEach(unit=>{unit.hp=0;});
  state.phase='german';state.orderPhase='combat';state.weather='Overcast';
  const attackers=['card-g17','card-g14'].map(id=>state.units.find(unit=>unit.id===id));
  const defender=state.units.find(unit=>unit.id==='card-a1');
  Object.assign(attackers[0],point(5,0),{hp:114,readiness:7,fired:false,combatCount:0});
  Object.assign(attackers[1],point(6,1),{hp:106,readiness:7,fired:false,combatCount:0});
  Object.assign(defender,point(6,0),{hp:34,readiness:0,defended:false});
  for(const unit of [...attackers,defender]) state.map[game.key(unit.x,unit.y)].terrain='plain';
  state.artillery={german:0,american:0};
  let seed=2147483648;
  for(let roll=0;roll<6;roll++) seed=seedForRoll(seed);
  state.seed=seed;
  const ids=attackers.map(unit=>unit.id);
  assert.equal(rules.preview(state,ids,defender.id,{strategy:'major',defenseStrategy:'delay'}).odds,11);
  let result=rules.resolveAttack(state,ids,defender.id,{strategy:'major',defenseStrategy:'delay'});
  assert.equal(result.odds,11);
  assert.equal(result.attackStrategy,'major');
  assert.equal(result.defenseStrategy,'delay');
  assert.deepEqual(result.units.filter(unit=>unit.side==='german').map(unit=>unit.loss),[0,0]);
  assert.equal(result.units.find(unit=>unit.id===defender.id).loss,13);
  assert.equal(state.pendingBattle.retreatCount,3);
  while(state.pendingBattle) {
    const task=rules.maneuvers.current(state);
    const choice=rules.maneuvers.choices(state).find(choice=>
      !game.neighbors(state,choice.tile).some(tile=>game.occupant(state,tile)?.side==='german'));
    result=rules.chooseManeuver(state,task.kind==='retreat'?choice.direction:'finish');
  }
  assert.equal(result.retreatPath.length,3);
  assert.equal(result.units.find(unit=>unit.id===defender.id).loss,13);
  assert.deepEqual(state.casualties,{german:0,american:13});
});

test('combat calibration survives saves while older campaigns retain their casualty formula', () => {
  const {state,attacker,defender}=advanceFixture(1,2147483648);
  attacker.fired=false;attacker.combatCount=0;defender.hp=75;
  assert.equal(state.combatVersion,1);
  const revised=restore(state),legacy=structuredClone(state);
  delete legacy.combatVersion;
  const restoredLegacy=restore(legacy);
  assert.equal(restoredLegacy.combatVersion,undefined);
  const options={strategy:'major',defenseStrategy:'delay'};
  const revisedResult=rules.resolveAttack(revised,[attacker.id],defender.id,options);
  const legacyResult=rules.resolveAttack(restoredLegacy,[attacker.id],defender.id,options);
  assert.ok(revisedResult.units.find(unit=>unit.id===defender.id).loss<legacyResult.units.find(unit=>unit.id===defender.id).loss);
  const invalid=structuredClone(state);invalid.combatVersion=2;
  assert.throws(()=>game.deserialize(JSON.stringify(invalid)),/combat rules revision/);
  delete invalid.appleSupplementVersion;
  invalid.combatVersion=1;
  assert.throws(()=>game.deserialize(JSON.stringify(invalid)),/supplement/i);
});

test('resource-scale supply previews agree with delivery probabilities and never exceed 100 percent', () => {
  const {state,attacker,defender}=advanceFixture(1,0);
  attacker.fired=false;attacker.combatCount=0;defender.hp=75;
  state.generalSupply.german=789;
  assert.match(rules.preview(state,[attacker.id],defender.id).risk,/attacker 50%, defender 100%/);
  state.generalSupply.german=2000;
  assert.match(rules.preview(state,[attacker.id],defender.id).risk,/attacker 100%, defender 100%/);
  restore(state);
});

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
    s=>{s.airInterdiction.fuel=5*rules.provisional.interdictionFuelScale+1;},
    s=>{s.airInterdiction.supply=5*rules.provisional.interdictionSupplyScale+1;},
    s=>{s.airInterdiction.movement.unknown=1;},
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

test('rule 17.1 lifts the British Meuse ban once any German unit has exited', () => {
  const state=britishFixture();
  const unit=state.units.find(unit=>unit.nationality==='british');
  state.units.forEach(unit=>{unit.hp=0;});
  Object.assign(unit,point(1,4),{hp:unit.maxHp});
  const destination=state.map[game.key(point(2,4).x,4)];
  const key=game.key(destination.x,destination.y);
  unit.mp=rules.movementCost(state,unit,destination,unit);
  assert.equal(game.referenceMap.crossesMeuse(unit,destination),true);
  assert.equal(game.reachable(state,unit).has(key),false);
  state.exitPoints=0.25;
  assert.equal(game.reachable(state,unit).has(key),true);
  state.exitPoints=0;
  assert.equal(game.reachable(state,unit).has(key),false);
  const german=state.units.find(unit=>unit.side==='german');
  german.exited=true;
  assert.equal(game.reachable(state,unit).has(key),true);
  assert.equal(game.move(state,unit.id,destination.x,destination.y),true);
  assert.equal(game.key(unit.x,unit.y),key);
});

test('calibrated attacker losses stay one tenth of exposure across odds', () => {
  function battle(attackerHp,defenderHp,roll) {
    const state=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
    state.units.forEach(unit=>{unit.hp=0;});
    state.phase='german';state.orderPhase='combat';state.weather='Overcast';
    state.artillery={german:0,american:0};
    const attacker=state.units.find(unit=>unit.id==='card-g1');
    const defender=state.units.find(unit=>unit.id==='card-a1');
    Object.assign(attacker,point(5,0),{hp:attackerHp,maxHp:attackerHp,readiness:0,fired:false,combatCount:0});
    Object.assign(defender,point(6,0),{hp:defenderHp,maxHp:defenderHp,readiness:0,defended:false,combatCount:0});
    for(const tile of [point(5,0),point(6,0)]) state.map[game.key(tile.x,tile.y)].terrain='plain';
    let seed=roll;
    for(let step=0;step<5;step++) seed=seedForRoll(seed);
    state.seed=seed;
    return {state,attacker,defender};
  }
  for(const [attackerHp,defenderHp,roll,attackLoss,defenseLoss] of [
    [80,80,0,2,1],[80,80,2147483648,2,2],[80,80,4294967295,1,3],
    [40,80,0,4,1],[40,80,2147483648,4,1],[40,80,4294967295,3,2],
    [80,40,2147483648,1,3]
  ]) {
    const {state,attacker,defender}=battle(attackerHp,defenderHp,roll);
    assert.equal(rules.preview(state,[attacker.id],defender.id,{strategy:'assault',defenseStrategy:'hold'}).odds,attackerHp/defenderHp);
    const result=rules.resolveAttack(state,[attacker.id],defender.id,{strategy:'assault',defenseStrategy:'hold',artillery:0});
    assert.equal(result.attackStrategy,'assault');
    assert.equal(result.defenseStrategy,'hold');
    assert.equal(result.units.find(unit=>unit.id===attacker.id).loss,attackLoss);
    assert.equal(result.units.find(unit=>unit.id===defender.id).loss,defenseLoss);
  }
});

test('artillery caps follow the side that owns them', () => {
  assert.deepEqual([rules.artilleryMaximum(game.createGame(42,'manual'),'german'),rules.artilleryMaximum(game.createGame(42,'manual'),'american')],[8,5]);
  const legacy=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
  delete legacy.resourceVersion;
  assert.deepEqual([rules.artilleryMaximum(legacy,'german'),rules.artilleryMaximum(legacy,'american')],[6,4]);
  const german=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
  german.defensiveArtillery=8;
  restore(german);
  german.defensiveArtillery=9;
  assert.throws(()=>game.deserialize(JSON.stringify(german)),/Invalid/);
  const allied=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
  allied.playerSide='american';allied.phase='american';allied.orderPhase='movement';
  allied.defensiveArtillery=5;
  restore(allied);
  allied.defensiveArtillery=6;
  assert.throws(()=>game.deserialize(JSON.stringify(allied)),/Invalid/);
  legacy.playerSide='american';legacy.phase='american';legacy.orderPhase='movement';
  legacy.generalSupply={german:120,american:120};legacy.fuel=120;legacy.artillery={german:6,american:4};
  legacy.defensiveArtillery=4;
  const restoredLegacy=restore(legacy);
  restoredLegacy.defensiveArtillery=5;
  assert.throws(()=>game.deserialize(JSON.stringify(restoredLegacy)),/Invalid/);
  const hotseat=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
  hotseat.defensiveArtillery=8;
  rules.beginSide(hotseat,'german');
  assert.equal(hotseat.defensiveArtillery,8);
  rules.beginSide(hotseat,'american');
  assert.equal(hotseat.playerSide,'american');
  assert.equal(hotseat.defensiveArtillery,5);

  function started(attackerId,defenderId) {
    const state=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
    state.units.forEach(unit=>{unit.hp=0;});
    state.orderPhase='combat';
    const attacker=state.units.find(unit=>unit.id===attackerId);
    const defender=state.units.find(unit=>unit.id===defenderId);
    Object.assign(attacker,point(5,0),{hp:attacker.maxHp,fired:false,combatCount:0});
    Object.assign(defender,point(6,0),{hp:defender.maxHp,defended:false,combatCount:0});
    state.phase=attacker.side;state.playerSide=attacker.side;
    const result=rules.maneuvers.start(state,{ids:[attacker.id],defenderId:defender.id,
      before:[attacker,defender].map(unit=>({id:unit.id,name:unit.name,side:unit.side,hp:unit.hp})),
      original:{x:defender.x,y:defender.y},attackerPositions:[{id:attacker.id,x:attacker.x,y:attacker.y}],
      odds:2,attackStrategy:'assault',defenseStrategy:'hold',artillery:0,defensiveArtillery:0,
      events:[],retreatCount:1,advanceLimit:0,attackerMustRetreat:false});
    assert.equal(result.pending,true);
    return state;
  }
  const germanAttack=started('card-g1','card-a1');
  germanAttack.pendingBattle.artillery=8;
  germanAttack.pendingBattle.defensiveArtillery=5;
  restore(germanAttack);
  germanAttack.pendingBattle.defensiveArtillery=6;
  assert.throws(()=>game.deserialize(JSON.stringify(germanAttack)),/maneuver/);
  germanAttack.pendingBattle.defensiveArtillery=5;
  germanAttack.pendingBattle.artillery=9;
  assert.throws(()=>game.deserialize(JSON.stringify(germanAttack)),/maneuver/);
  const alliedAttack=started('card-a1','card-g1');
  alliedAttack.pendingBattle.artillery=5;
  alliedAttack.pendingBattle.defensiveArtillery=8;
  restore(alliedAttack);
  alliedAttack.pendingBattle.artillery=6;
  assert.throws(()=>game.deserialize(JSON.stringify(alliedAttack)),/maneuver/);
  alliedAttack.pendingBattle.artillery=5;
  alliedAttack.pendingBattle.defensiveArtillery=9;
  assert.throws(()=>game.deserialize(JSON.stringify(alliedAttack)),/maneuver/);
  const html=require('node:fs').readFileSync(require('node:path').join(__dirname,'index.html'),'utf8');
  assert.match(html,/id="battle-defense-artillery"[^>]*max="8"/);
  assert.match(html,/id="defensive-artillery"[^>]*max="8"/);
});

test('example-scale interdiction, refill and fuel use are multiples of the old 120-point steps', () => {
  const {interdictionFuelScale,interdictionSupplyScale,refill,fuelMultiplier}=rules.provisional;
  assert.equal(rules.armorFuel(game.createGame(1,'manual'),'plain'),fuelMultiplier*Math.ceil(rules.COST.plain/2));
  assert.equal(rules.armorFuel(game.createGame(1,'manual'),'attack'),fuelMultiplier*2);
  const legacy=game.createGame(1,'manual','german',2,{playMode:'hotseat'});
  delete legacy.resourceVersion;
  assert.equal(rules.armorFuel(legacy,'plain'),Math.ceil(rules.COST.plain/2));
  assert.equal(rules.armorFuel(legacy,'attack'),2);
  const state=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
  state.turn=8;state.weather='Clear';state.airInterdiction.turn=1;
  state.seed=seedForRoll(Math.floor(0.99*4294967296));
  const fuel=state.fuel,supply=state.generalSupply.german;
  rules.beginSide(state,'german');
  assert.equal(state.airInterdiction.fuel,Math.min(fuel,5*interdictionFuelScale));
  assert.equal(state.airInterdiction.supply,Math.min(supply,Math.floor((state.airInterdiction.supply/interdictionSupplyScale))*interdictionSupplyScale));
  assert.ok(state.airInterdiction.fuel>5);
  assert.equal(state.airInterdiction.fuel%interdictionFuelScale,0);
  assert.equal(state.airInterdiction.supply%interdictionSupplyScale,0);
  assert.ok(state.airInterdiction.supply<=5*interdictionSupplyScale);
  state.weather='Overcast';state.turn=9;state.seed=seedForRoll(1);
  const before=state.seed;
  rules.beginSide(state,'german');
  assert.equal(state.airInterdiction.fuel,0);
  assert.equal(state.airInterdiction.supply,0);
  assert.notEqual(state.seed,before);
  legacy.turn=8;legacy.weather='Clear';legacy.fuel=120;
  legacy.generalSupply={german:120,american:120};legacy.artillery={german:6,american:4};
  legacy.airInterdiction.turn=1;legacy.seed=seedForRoll(Math.floor(0.99*4294967296));
  rules.beginSide(legacy,'german');
  assert.ok(legacy.airInterdiction.fuel<=5&&legacy.airInterdiction.supply<=5);
  const moved=game.createGame(7,'manual','german',2,{playMode:'hotseat'});
  moved.turn=2;moved.phase='german';moved.orderPhase='movement';
  moved.units.forEach(unit=>{unit.hp=0;});
  const armor=moved.units.find(unit=>unit.side==='german'&&unit.type==='armor');
  const origin=point(8,5),destination=moved.map[game.key(point(9,5).x,5)];
  assert.equal(game.referenceMap.crossesRiver(origin,destination),false);
  Object.assign(armor,origin,{hp:armor.maxHp,mp:30});
  const pool=moved.fuel;
  assert.equal(game.move(moved,armor.id,destination.x,destination.y),true);
  assert.equal(pool-moved.fuel,rules.armorFuel(moved,destination.terrain));
  assert.equal(refill.german,70);
  assert.equal(refill.american,36);
});

test('rule 16.2 prints BRIDGE BLOWN and stops a German river crossing', () => {
  function crossing(roll,options) {
    const state=game.createGame(42,'manual','german',2,{playMode:'hotseat',...options});
    state.turn=2;state.phase='german';state.orderPhase='movement';state.weather='Overcast';
    state.units.forEach(unit=>{unit.hp=0;});
    const unit=state.units.find(unit=>unit.side==='german'&&unit.type==='armor');
    const edge=game.referenceMap.riverEdges.find(({a,b})=>state.map[game.key(a.x,a.y)]&&state.map[game.key(b.x,b.y)]);
    Object.assign(unit,edge.a,{hp:unit.maxHp,mp:30,exited:false});
    state.seed=seedForRoll(Math.floor(roll*4294967296));
    const fuel=state.fuel;
    const moved=game.move(state,unit.id,edge.b.x,edge.b.y);
    return {state,unit,moved,fuel,to:edge.b};
  }
  const blown=crossing(0.05);
  assert.equal(blown.moved,false);
  assert.match(blown.state.log[0],/BRIDGE BLOWN/);
  assert.equal(blown.unit.mp,0);
  assert.equal(blown.state.fuel,blown.fuel);
  assert.notEqual(game.key(blown.unit.x,blown.unit.y),game.key(blown.to.x,blown.to.y));
  const crossed=crossing(0.8);
  assert.equal(crossed.moved,true);
  assert.equal(game.key(crossed.unit.x,crossed.unit.y),game.key(crossed.to.x,crossed.to.y));
  assert.equal(crossed.fuel-crossed.state.fuel,rules.armorFuel(crossed.state,crossed.state.map[game.key(crossed.to.x,crossed.to.y)].terrain));
  assert.ok(!crossed.state.log.some(line=>line.includes('BRIDGE BLOWN')));
  const older=crossing(0.05,{appleSupplement:false});
  assert.equal(older.moved,true);
  assert.ok(!older.state.log.some(line=>line.includes('BRIDGE BLOWN')));
});

test('rule 5.8 fogs enemy CP for display and leaves true strength in combat', () => {
  const state=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
  state.playMode='human';state.phase='german';state.playerSide='german';
  const enemy=state.units.find(unit=>unit.side==='american'&&unit.hp>0);
  const friend=state.units.find(unit=>unit.side==='german'&&unit.hp>0);
  const seed=state.seed,trueHp=enemy.hp;
  const shown=rules.displayedCombatPoints(state,enemy);
  assert.equal(state.seed,seed);
  assert.equal(enemy.hp,trueHp);
  assert.ok(shown>=trueHp+1&&shown<=trueHp+rules.provisional.fog.human);
  assert.equal(rules.displayedCombatPoints(state,enemy),shown);
  assert.equal(rules.displayedCombatPoints(state,friend),friend.hp);
  state.phase='american';
  const coarse=rules.displayedCombatPoints(state,enemy);
  assert.ok(coarse>=trueHp+1&&coarse<=trueHp+rules.provisional.fog.computer);
  assert.equal(rules.enemyStrengthOffset(state,friend),0);
  state.playMode='computer';
  assert.equal(rules.displayedCombatPoints(state,enemy),trueHp);
  function battle(seedValue) {
    const fixture=game.createGame(42,'manual','german',2,{playMode:'hotseat'});
    fixture.units.forEach(unit=>{unit.hp=0;});
    fixture.phase='german';fixture.orderPhase='combat';fixture.weather='Overcast';
    fixture.playMode='human';fixture.playerSide='german';
    fixture.artillery={german:0,american:0};
    const attacker=fixture.units.find(unit=>unit.id==='card-g1');
    const defender=fixture.units.find(unit=>unit.id==='card-a1');
    Object.assign(attacker,point(5,0),{hp:80,maxHp:80,readiness:0,fired:false,combatCount:0});
    Object.assign(defender,point(6,0),{hp:80,maxHp:80,readiness:0,defended:false,combatCount:0});
    for(const tile of [point(5,0),point(6,0)]) fixture.map[game.key(tile.x,tile.y)].terrain='plain';
    fixture.seed=seedValue;
    return {fixture,attacker,defender};
  }
  const roll=seedForRoll(2147483648);
  const fogged=battle(roll),plain=battle(roll);
  rules.displayedCombatPoints(fogged.fixture,fogged.defender);
  assert.ok(rules.displayedCombatPoints(fogged.fixture,fogged.defender)>fogged.defender.hp);
  const fogResult=rules.resolveAttack(fogged.fixture,[fogged.attacker.id],fogged.defender.id,{strategy:'assault',defenseStrategy:'hold',artillery:0});
  const plainResult=rules.resolveAttack(plain.fixture,[plain.attacker.id],plain.defender.id,{strategy:'assault',defenseStrategy:'hold',artillery:0});
  assert.equal(fogResult.odds,plainResult.odds);
  assert.deepEqual(fogResult.units.map(unit=>unit.loss),plainResult.units.map(unit=>unit.loss));
});

test('rule 10.5 airdrops succeed more often in Clear weather than Scattered', () => {
  assert.ok(rules.provisional.airdrop.Clear>rules.provisional.airdrop.Scattered);
  function attempt(weather,roll) {
    const state=game.createGame(1,'manual','german',2,{playMode:'hotseat'});
    state.units.forEach(unit=>{unit.hp=0;});
    const unit=state.units.find(unit=>unit.side==='american');
    Object.assign(unit,point(8,5),{hp:40,maxHp:80,exited:false});
    const deltas=[[-1,-1],[1,-1],[2,0],[1,1],[-1,1],[-2,0]];
    state.units.filter(unit=>unit.side==='german').slice(0,6).forEach((enemy,index)=>{
      const [dx,dy]=deltas[index];
      Object.assign(enemy,{hp:30,maxHp:50,exited:false,x:unit.x+dx,y:unit.y+dy});
    });
    assert.equal(rules.localSupply(state,unit),'isolated');
    state.turn=6;state.weather=weather;
    state.seed=seedForRoll(Math.floor(roll*4294967296));
    rules.attemptAirdrops(state);
    return unit.airdropTurn;
  }
  const between=(rules.provisional.airdrop.Scattered+rules.provisional.airdrop.Clear)/2;
  assert.equal(attempt('Clear',between),6);
  assert.equal(attempt('Scattered',between),undefined);
  assert.equal(attempt('Overcast',0.01),undefined);
  const state=game.createGame(1,'manual','german',2,{playMode:'hotseat'});
  state.weather='Cloudy';state.turn=6;const seed=state.seed;
  rules.attemptAirdrops(state);
  assert.equal(state.seed,seed);
});

test('German AI takes Bastogne and the key towns before exits and stacks good-odds attacks', () => {
  const state=game.createGame(1,'manual','german',2,{playMode:'computer'});
  state.turn=2;state.phase='german';state.orderPhase='movement';
  const towns=Object.values(state.map).filter(tile=>tile.sourcePoints);
  let towardTown=0,orders=0;
  for(const unit of state.units.filter(unit=>unit.side==='german'&&unit.hp>0)) {
    unit.mp=unit.mobility;
    const order=game.planMovement(state,unit);
    if(!order) continue;
    orders++;
    assert.equal(order.tile.exit,false);
    if(towns.some(town=>game.distance(order.tile,town)<game.distance(unit,town))) towardTown++;
  }
  assert.ok(orders>0&&towardTown>0);
  const early=game.createGame(3,'manual','german',2,{playMode:'computer'});
  early.turn=2;early.phase='german';early.orderPhase='movement';
  early.units.forEach(unit=>{unit.hp=0;});
  const holders=early.units.filter(unit=>unit.side==='german');
  towns.forEach((tile,index)=>Object.assign(holders[index],{hp:40,maxHp:40,exited:false,x:tile.x,y:tile.y}));
  const rusher=holders[3];
  Object.assign(rusher,point(16,4),{hp:100,maxHp:100,mp:30,exited:false});
  const earlyOrder=game.planMovement(early,rusher);
  assert.ok(!earlyOrder||earlyOrder.tile.exit===false);
  const late=game.createGame(3,'manual','german',2,{playMode:'computer'});
  late.turn=8;late.phase='german';late.orderPhase='movement';
  late.units.forEach(unit=>{unit.hp=0;});
  const lateHolders=late.units.filter(unit=>unit.side==='german');
  Object.values(late.map).filter(tile=>tile.sourcePoints).forEach((tile,index)=>{
    Object.assign(lateHolders[index],{hp:40,maxHp:40,exited:false,x:tile.x,y:tile.y});
    tile.owner='german';
  });
  const spear=lateHolders.find(unit=>unit.hp===0&&unit.type==='armor');
  Object.assign(spear,point(4,2),{hp:100,maxHp:100,mp:30,exited:false});
  const exitOrder=game.planMovement(late,spear);
  assert.ok(exitOrder&&exitOrder.tile.exit);

  const fight=game.createGame(42,'manual','german',2,{playMode:'computer'});
  fight.units.forEach(unit=>{unit.hp=0;});
  fight.phase='german';fight.orderPhase='combat';fight.playMode='computer';
  fight.artillery={german:8,american:0};fight.weather='Overcast';
  const bastogne=Object.values(fight.map).find(tile=>tile.name==='BASTOGNE');
  const adjacent=game.neighbors(fight,bastogne).filter(tile=>!game.referenceMap.crossesRiver(tile,bastogne)).slice(0,3);
  const germans=fight.units.filter(unit=>unit.side==='german');
  adjacent.forEach((tile,index)=>Object.assign(germans[index],{hp:100,maxHp:100,x:tile.x,y:tile.y,fired:false,combatCount:0,exited:false,readiness:6}));
  const garrison=fight.units.find(unit=>unit.side==='american');
  Object.assign(garrison,{hp:70,maxHp:70,x:bastogne.x,y:bastogne.y,defended:false,exited:false,readiness:6,fired:false,combatCount:0});
  const battle=rules.nextBattle(fight,'german');
  assert.equal(battle.defender.id,garrison.id);
  assert.ok(battle.ids.length>=3);
  assert.ok(battle.odds>=1.15);
  const plain=game.createGame(42,'manual','german',2,{playMode:'computer'});
  plain.units.forEach(unit=>{unit.hp=0;});
  plain.phase='german';plain.orderPhase='combat';plain.playMode='computer';
  plain.artillery={german:0,american:0};plain.weather='Overcast';
  const attacker=plain.units.find(unit=>unit.id==='card-g1');
  const defender=plain.units.find(unit=>unit.id==='card-a1');
  Object.assign(attacker,point(5,0),{hp:115,maxHp:115,readiness:6,fired:false,combatCount:0});
  Object.assign(defender,point(6,0),{hp:20,maxHp:20,readiness:0,defended:false,combatCount:0});
  for(const tile of [point(5,0),point(6,0)]) plain.map[game.key(tile.x,tile.y)].terrain='plain';
  const major=rules.planBattle(plain,attacker);
  assert.equal(major.options.strategy,'major');
  assert.ok(major.odds>=2.5);
});
