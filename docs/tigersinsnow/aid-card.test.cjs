'use strict';
const assert = require('node:assert/strict');
const test = require('node:test');
const game = require('./game.js');
const card = game.aidCard;
const point = (col,row) => game.referenceMap.position(col,row);
const tile = (state,col,row) => state.map[game.key(point(col,row).x,row)];
const restore = state => {
  const copy = game.deserialize(JSON.stringify(state));
  assert.deepEqual(copy,state);
  return copy;
};

test('unit presentation distinguishes airborne and mechanised formations without changing rule types', () => {
  for (const [name, type, symbol] of [
    ['82nd AIRB DIV', 'infantry', 'P'], ['3rd FJ DIV', 'infantry', 'P'],
    ['101 Airborne', 'infantry', 'P'], ['3rd PZGR DIV', 'armor', 'M'],
    ['14th CALV REG', 'armor', 'M'], ['2nd PZ DIV', 'armor', '='],
    ['99th INF DIV', 'infantry', 'X']
  ]) {
    const unit = { name, type };
    assert.equal(game.unitPresentation(unit).symbol, symbol);
    assert.deepEqual(unit, { name, type });
  }
  assert.equal(card.roster.filter(unit => game.unitPresentation(unit).symbol === 'P').length, 4);
  assert.equal(card.roster.filter(unit => game.unitPresentation(unit).symbol === 'M').length, 3);
});

test('aid card includes every formation, unaggregated strength and correct mechanized classification', () => {
  assert.equal(card.roster.length,66);
  assert.equal(new Set(card.roster.map(unit => unit.id)).size,66);
  const expected = {
    german: [55,50,50,55,50,35,50,65,65,55,65,70,70,100,100,100,115,115,75,45,115,115,50,60,75,100,45,55],
    american: [80,75,75,55,20,20,30,75,45,35,35,30,40,40,40,40,40,40,70,70,75,70,100,100,65,45,70,70,70,55,100,100,65,35,70,65,70,70]
  };
  for (const side of ['german','american']) assert.deepEqual(card.roster.filter(unit => unit.side === side).map(unit => unit.hp),expected[side]);
  assert.equal(card.roster.filter(unit => unit.turn === 0).length,29);
  assert.equal(card.roster.filter(unit => unit.turn > 0).length,37);
  assert.equal(card.roster.filter(unit => unit.nationality === 'british').length,6);
  assert.ok(card.roster.filter(unit => /PZGR|CALV/.test(unit.name)).every(unit => unit.type === 'armor'));
  assert.ok(!card.roster.some(unit => /Garrison/.test(unit.name)));
  assert.equal(card.roster.find(unit => unit.name === '12th SS PZ DIV').turn,1);
  assert.deepEqual(card.directions,[[1,-1],[2,0],[1,1],[-1,1],[-2,0],[-1,-1]]);
  assert.equal(Object.keys(card.abbreviations).length,19);
});

test('all 37 dated entries match the aid card and the rebuilt map', () => {
  const expected = [
    ['card-g18',16,13],['card-g19',16,12],['card-g20',16,11],
    ['card-g21',18,14],['card-g22',18,13],['card-g23',18,12],
    ['card-g24',22,16],['card-g25',22,14],['card-g26',22,15],['card-g27',22,13],['card-g28',24,15],
    ['card-a13',16,18],['card-a14',16,20],['card-a15',16,17],['card-a16',16,10],['card-a17',16,9],['card-a18',16,8],
    ['card-a19',16,28],['card-a20',16,24],['card-a21',17,10],['card-a22',19,8],['card-a23',20,6],['card-a24',20,21],
    ['card-a25',20,7],['card-a26',21,26],['card-a27',21,20],['card-a28',21,19],['card-a29',21,5],['card-a30',21,22],
    ['card-a31',22,2],['card-a32',22,1],['card-a33',22,27],['card-a34',22,3],['card-a35',23,25],['card-a36',24,4],
    ['card-a37',25,23],['card-a38',25,6]
  ];
  assert.deepEqual(card.roster.filter(unit => unit.turn).map(unit => [unit.id,unit.turn+15,unit.entry]),expected);
  const state = game.createGame(42,'manual');
  assert.equal(state.units.length,29);
  assert.equal(game.manual.offMapReinforcements(state).length,37);
  for (let day=16;day<=25;day++) {
    state.turn=day-15;
    for (const unit of state.units) unit.hp=0;
    game.manual.reinforce(state);
    for (const entry of card.roster.filter(unit => unit.turn === state.turn)) {
      const location = game.referenceMap.entries.find(location => location.number === entry.entry);
      const unit = state.units.find(unit => unit.id === entry.id);
      assert.equal(unit.hp,entry.hp);
      assert.equal(game.key(unit.x,unit.y),game.key(location.x,location.y));
    }
    restore(state);
  }
  assert.equal(state.units.length,66);
});

test('rating levels change every starting and inbound formation exactly once and survive saves', () => {
  for (let rating=1;rating<=9;rating++) {
    const state=game.createGame(42,'manual','german',2,{ratings:{german:rating,american:rating}});
    for(const unit of state.units) assert.equal(unit.maxHp,card.roster.find(entry=>entry.id===unit.id).hp+3*(rating-5));
    for(const entry of game.manual.offMapReinforcements(state)) {
      assert.equal(entry.strength,card.roster.find(unit=>unit.id===entry.id).hp+3*(rating-5));
    }
    restore(restore(state));
  }
  for(const ratings of [{german:0,american:5},{german:10,american:5},{german:5.5,american:5},{german:5,american:'5'}]) {
    assert.throws(()=>game.createGame(42,'manual','german',2,{ratings}),/rating/i);
  }
});

test('card edge and town VP require actual occupation and accumulate separately for each side', () => {
  for(const computer of [false,true]) {
    const state=game.createGame(42,'manual','german',2,computer?{playMode:'computer'}:{});
    state.units.forEach(unit=>{unit.hp=0;});
    const german=state.units.filter(unit=>unit.side==='german'),allied=state.units.filter(unit=>unit.side==='american');
    [[0,8],[2,10],[10,0],[7,4],[6,5],[11,7]].forEach(([col,row],i)=>Object.assign(german[i],point(col,row),{hp:30}));
    Object.assign(allied[0],point(18,4),{hp:30});
    Object.assign(allied[1],point(18,3),{hp:30});
    const factor=computer?2:1;
    for(let i=0;i<2;i++) {
      card.assessDay(state,game);
      assert.equal(state.victoryPoints.german,(i+1)*(3+3+10+25+25+50)*factor);
      assert.equal(state.victoryPoints.american,(i+1)*20);
    }
    german.forEach(unit=>{unit.hp=0;});
    const before=state.victoryPoints.german;
    Object.values(state.map).filter(tile=>tile.value).forEach(tile=>{tile.owner='german';});
    card.assessDay(state,game);
    assert.equal(state.victoryPoints.german,before,'Empty owned towns cannot generate points');
  }
  const state=game.createGame(42,'manual','american');
  assert.equal(card.multiplier(state),2);
});

test('scoring locations include only marked map hexes', () => {
  const state=game.createGame(42,'manual');
  assert.equal(card.edgeAward(tile(state,0,8),'german'),3);
  assert.equal(card.edgeAward(tile(state,0,9),'german'),3);
  assert.equal(card.edgeAward(tile(state,13,10),'german'),3);
  assert.equal(card.edgeAward(tile(state,14,10),'german'),0);
  assert.equal(card.edgeAward(tile(state,9,0),'german'),0);
  assert.equal(card.edgeAward(tile(state,10,0),'german'),10);
  assert.equal(card.edgeAward(tile(state,13,0),'german'),10);
  assert.equal(card.edgeAward(tile(state,14,0),'german'),0);
  assert.equal(card.edgeAward(tile(state,18,4),'american'),20);
  assert.equal(card.edgeAward(tile(state,18,6),'american'),20);
  assert.equal(card.edgeAward(tile(state,18,7),'american'),0);
});

test('card exits award remaining CP times (16 minus turn/4), exactly once and without a ZOC exception', () => {
  for(const turn of [1,7,12]) {
    const state=game.createGame(42,'manual');state.units.forEach(unit=>{unit.hp=0;});
    state.turn=turn;
    const unit=state.units[0],enemy=state.units.find(unit=>unit.side==='american');
    Object.assign(unit,point(0,4),{hp:55});
    Object.assign(enemy,point(1,4),{hp:50});
    card.assessDay(state,game);
    const points=55*(16-turn/4);
    assert.equal(state.exitPoints,points);
    assert.equal(unit.exited,true);
    assert.equal(game.occupant(state,tile(state,0,4)),undefined);
    assert.equal(state.victoryPoints.german,points);
    card.assessDay(state,game);
    assert.equal(state.exitPoints,points);
    restore(state);
  }
});

test('campaign and short-scenario victory boundaries reproduce every victory level', () => {
  const expected={
    campaign:[[50,'american','Strategic'],[51,'american','Operational'],[299,'american','Operational'],[300,'american','Marginal'],[699,'american','Marginal'],[700,'german','Marginal'],[949,'german','Marginal'],[950,'german','Operational'],[1199,'german','Operational'],[1200,'german','Strategic']],
    short:[[50,'american','Strategic'],[51,'american','Operational'],[174,'american','Operational'],[175,'american','Marginal'],[324,'american','Marginal'],[325,'german','Marginal'],[424,'german','Marginal'],[425,'german','Operational'],[699,'german','Operational'],[700,'german','Strategic']]
  };
  for(const scenario of ['campaign','short']) {
    const state=game.createGame(42,'manual','german',2,{scenario});
    for(const [points,leader,level] of expected[scenario]) {
      state.victoryPoints={german:points,american:0};
      assert.equal(game.campaignStanding(state).leader,leader);
      assert.equal(game.campaignStanding(state).level,level);
    }
    state.victoryPoints={german:0,american:100};
    assert.equal(game.campaignStanding(state).points,-100);
    assert.equal(game.campaignStanding(state).level,'Strategic');
    assert.equal(card.lastTurn(state),scenario==='short'?7:12);
    state.turn=card.lastTurn(state);state.units.forEach(unit=>{unit.hp=0;});
    game.manual.finishDay(state);
    assert.equal(state.winner,'american');
    restore(state);
  }
});

function battleFixture(side='german') {
  const state=game.createGame(123,'manual','german');
  state.playerSide=side;state.phase='german';state.orderPhase='combat';
  state.units.forEach(unit=>{unit.hp=0;});
  const attacker=state.units.find(unit=>unit.id==='card-g17'),defender=state.units.find(unit=>unit.id==='card-a2');
  Object.assign(attacker,point(5,0),{hp:115,fired:false});
  Object.assign(defender,point(6,0),{hp:1,defended:false});
  tile(state,5,0).terrain='plain';tile(state,6,0).terrain='plain';
  state.artillery={german:0,american:0};
  return {state,attacker,defender};
}

test('casualty VP equals actual CP removed, including elimination and maneuvers', () => {
  const {state,attacker,defender}=battleFixture();
  const before=state.units.map(unit=>({...unit}));
  game.manual.resolveAttack(state,[attacker.id],defender.id);
  while(state.pendingBattle) game.manual.chooseManeuver(state,'finish');
  const losses=Object.fromEntries(['german','american'].map(side=>[side,before.filter(unit=>unit.side===side)
    .reduce((sum,unit)=>sum+unit.hp-state.units.find(now=>now.id===unit.id).hp,0)]));
  assert.deepEqual(state.casualties,losses);
  assert.equal(state.victoryPoints.german,losses.american);
  assert.equal(state.victoryPoints.american,losses.german);
  assert.equal(state.score,losses.american-losses.german);
});

test('human advances allow F or only the vacated first hex and resume exactly after import', () => {
  for(const direction of ['finish',2]) {
    let {state,attacker,defender}=battleFixture();
    const result=game.manual.resolveAttack(state,[attacker.id],defender.id);
    assert.ok(result.pending);
    assert.equal(game.manual.maneuvers.current(state).kind,'advance');
    assert.deepEqual(game.manual.maneuvers.choices(state).map(choice=>choice.direction),[2]);
    assert.equal(game.endTurn(state,'german'),false);
    const before=JSON.stringify(state);
    assert.throws(()=>game.manual.chooseManeuver(state,1),/blocked/);
    assert.equal(JSON.stringify(state),before);
    // Map terrain is canonical on import; choose clear fixture hexes that are already clear.
    state.map=game.referenceMap.createMap();
    for(const tile of Object.values(state.map).filter(tile=>tile.sourcePoints)) tile.value=tile.sourcePoints;
    state=restore(state);
    const copy=restore(state);
    game.manual.chooseManeuver(state,direction);game.manual.chooseManeuver(copy,direction);
    while(state.pendingBattle) {
      game.manual.chooseManeuver(state,'finish');game.manual.chooseManeuver(copy,'finish');
    }
    assert.deepEqual(copy,state);
    assert.equal(state.pendingBattle,undefined);
    const unit=state.units.find(unit=>unit.id===attacker.id);
    assert.equal(unit.x,direction==='finish'?point(5,0).x:point(6,0).x);
  }
});

test('each computer attack pauses for human defensive strategy and artillery and survives imports', () => {
  let state=game.createGame(42,'manual','american');
  assert.ok(state.pendingDefense);
  assert.equal(state.pendingBattle,undefined);
  assert.equal(state.phase,'german');
  state=restore(state);
  const pending=structuredClone(state.pendingDefense),before=JSON.stringify(state);
  assert.throws(()=>game.manual.chooseDefense(state,'invalid',0),/valid/);
  assert.throws(()=>game.manual.chooseDefense(state,'hold',99),/valid/);
  assert.equal(JSON.stringify(state),before);
  assert.equal(game.endTurn(state,'american'),false);
  assert.equal(game.manual.eligibleAttackers(state,state.units.find(unit=>unit.id===pending.defenderId)).length,0);
  const copy=restore(state);
  const result=game.manual.chooseDefense(state,'hold',3);
  assert.deepEqual(game.manual.chooseDefense(copy,'hold',3),result);
  assert.equal(result.defensiveArtillery,3);
  assert.deepEqual(copy,state);
  restore(state);
  for(const mutate of [
    s=>{s.pendingDefense.ids=['unknown'];},
    s=>{s.pendingDefense.artillery=99;},
    s=>{s.aidFlow.side='american';}
  ]) {
    const corrupt=JSON.parse(before);mutate(corrupt);
    assert.throws(()=>game.deserialize(JSON.stringify(corrupt)),/Invalid/);
  }
});

test('human retreats accept directional steps or F with additional losses and casualty credit', () => {
  for(const stop of [true,false]) {
    let {state,attacker,defender}=battleFixture('american');
    defender.hp=75;
    state.aidFlow={side:'german',stage:'combat',index:17,after:'american'};
    const result=game.manual.resolveAttack(state,[attacker.id],defender.id,{defenseStrategy:'withdrawal'});
    assert.ok(result.pending);
    assert.equal(game.manual.maneuvers.current(state).kind,'retreat');
    const firstHp=state.units.find(unit=>unit.id===defender.id).hp;
    state.map=game.referenceMap.createMap();
    Object.values(state.map).filter(tile=>tile.sourcePoints).forEach(tile=>{tile.value=tile.sourcePoints;});
    state=restore(state);
    const choice=game.manual.maneuvers.choices(state)[0];
    assert.ok(choice);
    game.manual.chooseManeuver(state,stop?'finish':choice.direction);
    const now=state.units.find(unit=>unit.id===defender.id);
    if(stop) assert.ok(now.hp<firstHp);
    else assert.notEqual(game.key(now.x,now.y),game.key(defender.x,defender.y));
    while(state.pendingBattle) game.manual.chooseManeuver(state,'finish');
    assert.ok(state.casualties.american>=75-now.hp);
    assert.equal(state.phase,'american');
    restore(state);
  }
});

test('full-roster computer benchmark completes sixty seeds with cumulative scoring', t => {
  const wins={german:0,american:0};
  for(let seed=1;seed<=60;seed++) {
    const state=game.createGame(seed,'manual','german',2,{playMode:'computer'});
    for(let day=0;!state.winner&&day<12;day++) game.simulateDay(state);
    assert.equal(state.turn,12);assert.ok(state.winner);
    assert.equal(state.score,state.victoryPoints.german-state.victoryPoints.american);
    assert.equal(state.winner,card.standing(state).leader);
    restore(state);wins[state.winner]++;
  }
  t.diagnostic(`Full-roster player-aid benchmark: ${JSON.stringify(wins)}; ratings 5/5; adapted deployment, seeds 1-60`);
});

test('both human command sides finish full and short campaigns with saved maneuver decisions', () => {
  for(const side of ['german','american']) {
    for(const scenario of ['campaign','short']) {
      let state=game.createGame(42,'manual',side,2,{scenario});
      for(let i=0;!state.winner&&i<500;i++) {
        if(state.pendingDefense) {
          game.manual.chooseDefense(state,'hold',0);
        } else if(state.pendingBattle) {
          const choices=game.manual.maneuvers.choices(state);
          game.manual.chooseManeuver(state,choices[0]?.direction || 'finish');
        } else assert.ok(game.endTurn(state,side));
        const active=state.units.filter(unit=>unit.hp>0&&!unit.exited);
        assert.equal(new Set(active.map(unit=>game.key(unit.x,unit.y))).size,active.length);
        state=restore(state);
      }
      assert.ok(state.winner);
      assert.equal(state.turn,scenario==='short'?7:12);
      assert.equal(state.winner,game.campaignStanding(state).leader);
    }
  }
});

test('card spectator action stepping and whole-day simulation agree for full roster and ratings', () => {
  for(const scenario of ['campaign','short']) {
    let state=game.createGame(42,'manual','german',2,{scenario,playMode:'computer',ratings:{german:7,american:4}});
    const daily=structuredClone(state);
    let actions=0;
    while(!state.winner) {
      const action=game.simulateAction(state);assert.ok(action);actions++;
      if(action.dayEnd) {game.simulateDay(daily);assert.deepEqual(state,daily);}
      assert.equal(state.pendingBattle,undefined);
      state=restore(state);
      assert.ok(actions<2500);
    }
    assert.equal(state.turn,scenario==='short'?7:12);
    assert.equal(state.units.filter(unit=>unit.nationality==='british').length,scenario==='short'?5:6);
  }
});

test('two-human hotseat alternates command without AI actions and uses human German VP', () => {
  let state=game.createGame(42,'manual','german',2,{playMode:'hotseat',scenario:'short'});
  assert.equal(state.phase,'german');assert.equal(state.playerSide,'german');
  assert.equal(card.multiplier(state),1);
  assert.equal(state.orderPhase,'combat');
  assert.throws(()=>game.simulateDay(state),/Computer/);
  assert.ok(game.endTurn(state,'german'));
  assert.equal(state.phase,'american');assert.equal(state.playerSide,'american');
  assert.equal(card.multiplier(state),1);
  assert.equal(state.orderPhase,'movement');
  for(let i=0;!state.winner&&i<100;i++) {
    assert.ok(game.endTurn(state,state.playerSide));
    assert.equal(state.pendingDefense,undefined);
    state=restore(state);
  }
  assert.equal(state.turn,7);assert.ok(state.winner);
  assert.equal(game.simulateDay(state),false);
});

test('hotseat defender commitments validate before mutation and both humans control maneuvers', () => {
  let {state,attacker,defender}=battleFixture();
  state.playMode='hotseat';defender.hp=75;
  state.artillery={german:6,american:4};
  const before=JSON.stringify(state);
  assert.throws(()=>game.manual.resolveAttack(state,[attacker.id],defender.id,{defensiveArtillery:5}),/defensive artillery/);
  assert.equal(JSON.stringify(state),before);
  const preview=game.battlePreview(state,[attacker.id],defender.id,{defensiveArtillery:3,defenseStrategy:'withdrawal'});
  assert.ok(preview.odds>0);
  assert.equal(JSON.stringify(state),before);
  const result=game.manual.resolveAttack(state,[attacker.id],defender.id,{defensiveArtillery:3,defenseStrategy:'withdrawal'});
  assert.equal(result.defensiveArtillery,3);
  assert.ok(result.pending);
  assert.equal(game.manual.maneuvers.current(state).kind,'retreat');
  assert.equal(game.manual.maneuvers.current(state).id,defender.id);
  assert.equal(state.playerSide,'german');
  state.map=game.referenceMap.createMap();
  Object.values(state.map).filter(tile=>tile.sourcePoints).forEach(tile=>{tile.value=tile.sourcePoints;});
  state=restore(state);
  const copy=restore(state);
  const direction=game.manual.maneuvers.choices(state)[0].direction;
  game.manual.chooseManeuver(state,direction);
  game.manual.chooseManeuver(copy,direction);
  assert.deepEqual(state,copy);
  while(state.pendingBattle) {
    game.manual.chooseManeuver(state,'finish');
    game.manual.chooseManeuver(copy,'finish');
  }
  assert.deepEqual(state,copy);
  assert.equal(card.multiplier(state),1);
  restore(state);
});

test('card display options persist and finishing movement spends no resources', () => {
  for(const color of [true,false]) for(const hexes of [true,false]) {
    const state=game.createGame(42,'manual','german',2,{playMode:'hotseat',display:{color,hexes}});
    assert.deepEqual(state.display,{color,hexes});
    game.endTurn(state,'german');
    const unit=state.units.find(unit=>unit.side==='american'&&unit.mp>0);
    assert.ok(unit);
    const before=structuredClone(state);
    assert.equal(game.manual.finishMovement(state,unit.id),true);
    before.units.find(formation=>formation.id===unit.id).mp=0;
    assert.match(state.log[0],/movement finished/);
    assert.match(state.reports.american.at(-1).text,/ended its movement/);
    before.log=structuredClone(state.log);
    before.reports=structuredClone(state.reports);
    assert.deepEqual(state,before);
    assert.equal(game.manual.finishMovement(state,unit.id),false);
    restore(state);
  }
  assert.throws(()=>game.createGame(42,'classic','german',2,{playMode:'hotseat'}),/hotseat/);
  assert.throws(()=>game.createGame(42,'manual','german',1,{playMode:'hotseat'}),/hotseat/);
  assert.throws(()=>game.createGame(42,'manual','german',2,{playMode:'invalid'}),/player controls/);
});

test('new card settings and save schemas reject unknown revisions and corrupted data', () => {
  const state=game.createGame(42,'manual');
  for(const mutate of [
    s=>{s.aidCardVersion=2;},s=>{s.ratings.german=0;},s=>{s.scenario='invalid';},
    s=>{s.defensiveArtillery=7;},s=>{s.victoryPoints.german=-1;},s=>{s.victoryPoints.german=0.1;},
    s=>{s.units[0].nationality='british';},s=>{s.units[0].maxHp++;},s=>{s.units[0].beganInZoc='yes';},
    s=>{s.display.hexes='yes';},s=>{s.pendingBattle={};},s=>{s.aidFlow={};}
  ]) {
    const copy=structuredClone(state);mutate(copy);
    assert.throws(()=>game.deserialize(JSON.stringify(copy)),/Invalid|valid|rating|revision|settings/i);
  }
  assert.throws(()=>game.createGame(42,'manual','german',2,{scenario:'unknown'}),/scenario/i);
});
