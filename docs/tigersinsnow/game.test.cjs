'use strict';
const assert = require('node:assert/strict');
const test = require('node:test');
const engine = require('./game.js');
// Keep the existing regression fixtures on their original adapted rules revision.
const game = { ...engine, createGame: (seed, ruleset, side, revision) =>
  engine.createGame(seed, ruleset, side, revision, { aidCard: false }) };

test('reference map transcribes dimensions, terrain, towns, entries and shared river edges', () => {
  const state = game.createGame(42, 'manual'), reference = game.referenceMap;
  assert.equal(state.referenceVersion, 2);
  assert.deepEqual(game.mapDimensions(state), { cols: 19, rows: 11 });
  assert.equal(Object.keys(state.map).length, 209);
  assert.ok(reference.terrainRows.every(row => row.length === 19));
  const tile = (col, row) => {
    const point = reference.position(col, row);
    return state.map[game.key(point.x, point.y)];
  };
  assert.equal(tile(7, 4).name, 'MARCHE');
  assert.equal(tile(6, 5).name, 'ROCHEFORT');
  assert.equal(tile(11, 7).name, 'BASTOGNE');
  assert.equal(tile(16, 0).terrain, 'westwall');
  assert.equal(tile(18, 4).terrain, 'forest');
  assert.equal(tile(3, 6).terrain, 'plain');
  assert.equal(tile(14, 0).terrain, 'ridge');
  assert.deepEqual(reference.entries.map(entry => entry.number), Array.from({ length: 28 }, (_, i) => i + 1));
  for (const entry of reference.entries) assert.equal(tile(entry.col, entry.row).entry, entry.number);
  assert.deepEqual(reference.entries.find(entry => entry.number === 24), { number: 24, col: 6, row: 10, edge: 'south', x: 13, y: 10 });
  assert.equal(new Set(reference.riverEdges.map(({ a, b }) => reference.edgeKey(a, b))).size, reference.riverEdges.length);
  reference.riverEdges.forEach(({ a, b }, index) => {
    assert.ok(state.map[game.key(a.x, a.y)] && state.map[game.key(b.x, b.y)]);
    assert.equal(game.distance(a, b), 1);
    assert.ok(reference.crossesRiver(a, b) && reference.crossesRiver(b, a));
    const segment = reference.riverSegments[index];
    assert.ok(Math.abs(Math.hypot(segment[1][0] - segment[0][0], segment[1][1] - segment[0][1]) - 26) < 1e-8);
    for (const point of segment) {
      for (const hex of [a, b]) {
        const center = reference.center(hex);
        assert.ok(Array.from({ length: 6 }, (_, i) => (i * 60 - 30) * Math.PI / 180).some(angle =>
          Math.hypot(point[0] - center.x - 26 * Math.cos(angle), point[1] - center.y - 26 * Math.sin(angle)) < 1e-8));
      }
    }
  });
  assert.equal(tile(9, 0).exit, true);
  assert.equal(tile(10, 0).exit, false);
  assert.equal(tile(0, 8).exit, false);
});

test('reference river movement is seeded, symmetric and identical in planning and execution', () => {
  const reference = game.referenceMap, { a, b } = reference.riverEdges[0];
  const observed = { german: new Set(), american: new Set() };
  for (let seed = 0; seed < 50; seed++) {
    const state = game.createGame(seed, 'manual');
    for (const side of ['german', 'american']) {
      const before = JSON.stringify(state);
      const cost = game.manual.crossingCost(state, a, b, side);
      observed[side].add(cost);
      assert.ok(side === 'german' ? cost >= 3 && cost <= 7 : cost >= 2 && cost <= 5);
      assert.equal(game.manual.crossingCost(state, b, a, side), cost);
      assert.equal(JSON.stringify(state), before, 'Preview must not consume randomness');
    }
  }
  assert.deepEqual([...observed.german].sort(), [3, 4, 5, 6, 7]);
  assert.deepEqual([...observed.american].sort(), [2, 3, 4, 5]);
  for (const side of ['german', 'american']) {
    const state = game.createGame(42, 'manual');
    state.units.forEach(unit => { unit.hp = 0; });
    Object.values(state.map).forEach(tile => { tile.terrain = 'forest'; });
    const unit = state.units.find(unit => unit.side === side);
    Object.assign(unit, { ...a, hp: unit.maxHp, mp: 10 });
    state.phase = side; state.orderPhase = 'movement';
    state.map[game.key(b.x, b.y)].terrain = 'plain';
    const cost = 2 + game.manual.crossingCost(state, a, b, side);
    assert.equal(game.manual.movementCost(state, a, state.map[game.key(b.x, b.y)], unit), cost);
    assert.equal(game.reachable(state, unit).get(game.key(b.x, b.y)).remaining, 10 - cost);
    assert.ok(game.move(state, unit.id, b.x, b.y));
    assert.equal(unit.mp, 10 - cost);
    assert.equal(game.manual.crossingCost(state, b, b, side), 0);
  }
});

test('reference combat applies river penalties per attacker and side-specific Westwall modifiers', () => {
  const state = game.createGame(42, 'manual'), reference = game.referenceMap;
  state.units.forEach(unit => { unit.hp = 0; });
  const { a, b } = reference.riverEdges[0];
  const attacker = state.units.find(unit => unit.id === 'g1'), defender = state.units.find(unit => unit.id === 'a1');
  Object.assign(attacker, a, { hp: 100, readiness: 6 });
  Object.assign(defender, b, { hp: 75, readiness: 6 });
  state.map[game.key(a.x, a.y)].terrain = 'plain'; state.map[game.key(b.x, b.y)].terrain = 'plain';
  state.artillery = { german: 0, american: 0 };
  const crossed = game.combatOdds(state, attacker, defender);
  state.referenceVersion = 1;
  assert.equal(game.combatOdds(state, attacker, defender), crossed * 2);
  state.referenceVersion = 2;
  assert.equal(game.battlePreview(state, [attacker.id], defender.id).odds, crossed);
  const extra = state.units.find(unit => unit.id === 'g2');
  const dry = game.neighbors(state, defender).find(tile => !reference.crossesRiver(tile, defender));
  Object.assign(extra, { x: dry.x, y: dry.y, hp: 100, readiness: 6 });
  state.map[game.key(dry.x, dry.y)].terrain = 'plain';
  assert.equal(game.battlePreview(state, [attacker.id, extra.id], defender.id).odds,
    game.combatOdds(state, attacker, defender) + game.combatOdds(state, extra, defender));
  extra.hp = 0;
  const baseOdds = game.combatOdds(state, attacker, defender);
  state.map[game.key(a.x, a.y)].terrain = 'westwall';
  assert.equal(game.combatOdds(state, attacker, defender), baseOdds * 2);
  state.map[game.key(b.x, b.y)].terrain = 'westwall';
  assert.equal(game.manual.terrainDefense(state, state.map[game.key(b.x, b.y)], 'american'), 1);
  attacker.side = 'american'; defender.side = 'german';
  assert.equal(game.manual.terrainDefense(state, state.map[game.key(b.x, b.y)], 'german'), 2.75);
  state.map[game.key(a.x, a.y)].terrain = 'plain';
  const alliedPlain = game.combatOdds(state, attacker, defender);
  state.map[game.key(a.x, a.y)].terrain = 'westwall';
  assert.equal(game.combatOdds(state, attacker, defender), alliedPlain);
});

test('reference reinforcements use numbered entries and immediate appropriate-edge alternatives', () => {
  for (const [id, number, col, row] of [['a9',28,0,1], ['a10',21,2,10], ['a11',6,0,0], ['g9',14,18,0]]) {
    const state = game.createGame(42, 'manual');
    state.units = state.units.filter(unit => unit.id !== id);
    state.arrivals = state.arrivals.filter(arrival => arrival !== id);
    state.turn = 5;
    const entry = game.referenceMap.entries.find(entry => entry.number === number);
    const blocker = state.units.find(unit => unit.id === 'g1');
    blocker.x = entry.x; blocker.y = entry.y;
    game.manual.reinforce(state);
    const arrived = state.units.find(unit => unit.id === id);
    const expected = game.referenceMap.position(col, row);
    assert.deepEqual({ x: arrived.x, y: arrived.y }, expected);
    assert.ok(!state.pending.includes(id));
    assert.equal(game.manual.reinforcementSchedule(state).find(entry => entry.id === id).entry, number);
    assert.deepEqual(game.deserialize(JSON.stringify(state)), state);
  }
  const state = game.createGame(42, 'manual');
  for (const [id, number] of [['a8',24], ['a9',28]]) {
    const unit = state.units.find(unit => unit.id === id), entry = game.referenceMap.entries.find(entry => entry.number === number);
    assert.equal(game.key(unit.x, unit.y), game.key(entry.x, entry.y));
  }
});

test('reference reinforcements stay delayed when every appropriate alternate is occupied', () => {
  const state = game.createGame(42, 'manual');
  state.turn = 5;
  const alternatives = Object.values(state.map).filter(tile => tile.alternate === 'south');
  alternatives.forEach((tile, i) => {
    const occupied = game.occupant(state, tile);
    if (!occupied) state.units.push({ ...state.units[0], id: `block${i}`, x: tile.x, y: tile.y });
  });
  game.manual.reinforce(state);
  assert.ok(state.pending.includes('a10'));
  assert.ok(!state.units.some(unit => unit.id === 'a10'));
  assert.equal(game.manual.offMapReinforcements(state).find(entry => entry.id === 'a10').status, 'Delayed');
  game.occupant(state, alternatives[0]).hp = 0;
  game.manual.reinforce(state);
  assert.ok(!state.pending.includes('a10'));
  assert.ok(state.arrivals.includes('a10'));
});

test('all map revisions round-trip without migrating terrain, objectives or saved allowances', () => {
  for (const revision of [undefined, 1, 2]) {
    const state = game.createGame(42, 'manual', 'german', revision || 1);
    if (revision === undefined) delete state.referenceVersion;
    state.units.forEach(unit => { unit.mobility = unit.type === 'armor' ? 24 : 20; });
    const restored = game.deserialize(JSON.stringify(state));
    assert.deepEqual(restored, state);
    assert.equal(Object.keys(restored.map).length, revision === 2 ? 209 : 180);
    assert.deepEqual(game.mapDimensions(restored), revision === 2 ? { cols: 19, rows: 11 } : { cols: 15, rows: 12 });
  }
  const state = game.createGame(42, 'manual');
  for (const patch of [{ referenceVersion: 3 }, { initialSeed: undefined }, { units: state.units.map((unit, i) => i ? unit : { ...unit, x: 0, y: 0 }) }]) {
    assert.throws(() => game.deserialize(JSON.stringify({ ...state, ...patch })), /Invalid/i);
  }
  assert.equal(Object.keys(game.createGame(42).map).length, 180);
});

test('reference-map AI finishes sixty seeded campaigns with valid occupation and scoring', t => {
  const wins = { german: 0, american: 0 };
  for (let seed = 1; seed <= 60; seed++) {
    const state = game.createGame(seed, 'manual'); state.playMode = 'computer';
    for (let day = 0; !state.winner && day < 12; day++) game.simulateDay(state);
    assert.ok(state.winner);
    assert.equal(game.campaignStanding(state).target, 8);
    const active = state.units.filter(unit => unit.hp > 0 && !unit.exited);
    assert.equal(new Set(active.map(unit => game.key(unit.x, unit.y))).size, active.length);
    assert.ok(active.every(unit => state.map[game.key(unit.x, unit.y)]));
    wins[state.winner]++;
  }
  t.diagnostic(`reference map: ${JSON.stringify(wins)}; seeds 1-60; geographic fidelity is not a claim of scenario balance`);
});

test('reference-map German AI seeks northern exits after securing its town objectives', () => {
  const state = game.createGame(42, 'manual'), unit = state.units.find(unit => unit.id === 'g1');
  state.units.forEach(counter => { counter.hp = 0; });
  Object.values(state.map).forEach(tile => {
    tile.terrain = 'plain';
    if (tile.value) tile.owner = 'german';
  });
  Object.assign(unit, game.referenceMap.position(10, 0), { hp: 100, mp: 10 });
  state.orderPhase = 'movement';
  const order = game.planMovement(state, unit);
  assert.ok(order);
  assert.equal(order.tile.exit, true);
  assert.equal(order.tile.row, 0);
  assert.ok(game.move(state, unit.id, order.tile.x, order.tile.y));
  game.manual.finishDay(state);
  assert.equal(unit.exited, true);
  assert.equal(state.exitPoints, 10);
});

test('campaign history explains battle losses and event types without changing source logs', () => {
  const battle = game.describeDispatch('1 SS Panzer > 99 Infantry: 1.4:1 / losses 2:4.');
  assert.equal(battle.kind, 'Battle');
  assert.equal(battle.title, '1 SS Panzer attacked 99 Infantry');
  assert.match(battle.detail, /Attacker lost 2 strength; defender lost 4 strength/);
  assert.equal(game.describeDispatch('101 Airborne arrives.').kind, 'Reinforcement');
  assert.equal(game.describeDispatch('99 Infantry eliminated.').kind, 'Elimination');
  assert.equal(game.describeDispatch('99 Infantry retreats.').kind, 'Retreat');
  assert.equal(game.describeDispatch('2 Panzer captures BASTOGNE.').kind, 'Objective');
  assert.equal(game.describeDispatch('17 DEC / FOG / German orders.').title, '17 December — fog weather');
  assert.equal(game.describeDispatch('AMERICAN VICTORY / 0 objective points.').title, 'America won the campaign');
  assert.match(game.describeDispatch('g1 + g2 > a1: 2.0:1 / major vs hold.').title, /g1 and g2/);
  assert.equal(game.describeDispatch('An unfamiliar saved message.').title, 'An unfamiliar saved message.');
});

test('army situation summaries reflect strength, objectives and inactive formations in both modes', () => {
  for (const ruleset of ['classic', 'manual']) {
    const state = game.createGame(1, ruleset);
    for (const side of ['american', 'german']) {
      const active = state.units.filter(unit => unit.side === side && unit.hp > 0 && !unit.exited);
      const summary = game.armySummary(state, side);
      assert.equal(summary.formations, active.length);
      assert.equal(summary.strength, active.reduce((sum, unit) => sum + unit.hp, 0));
      assert.equal(summary.objectives.length, side === 'american' ? ruleset === 'manual' ? 3 : 5 : 0);
      active[0].hp = 0;
      assert.equal(game.armySummary(state, side).formations, active.length - 1);
      if (ruleset === 'manual') {
        active[1].exited = true;
        assert.equal(game.armySummary(state, side).formations, active.length - 2);
      }
    }
  }
});

test('balanced objective points transfer between sides without changing victory scoring', () => {
  for (const ruleset of ['classic', 'manual']) {
    const state = game.createGame(1, ruleset);
    const total = ruleset === 'manual' ? 7 : 15;
    assert.equal(game.armySummary(state, 'american').objectivePoints, total);
    assert.equal(game.armySummary(state, 'german').objectivePoints, 0);
    const town = Object.values(state.map).find(tile => tile.name === 'BASTOGNE');
    town.owner = 'german';
    assert.equal(game.armySummary(state, 'american').objectivePoints, total - 3);
    assert.equal(game.armySummary(state, 'german').objectivePoints, 3);
    const before = JSON.stringify(state);
    game.armySummary(state, 'american'); game.armySummary(state, 'german');
    assert.equal(JSON.stringify(state), before, 'Balance reporting must not modify victory state');
    if (ruleset === 'manual') {
      state.exitPoints = 10; state.alliedBreakthrough = 2;
      assert.equal(game.armySummary(state, 'german').objectivePoints, 3);
    }
  }
});

test('computer versus computer simulates both armies, saves mode and finishes in both rulesets', () => {
  for (const ruleset of ['classic', 'manual']) {
    for (const seed of [1, 42, 123]) {
      let state = game.createGame(seed, ruleset);
      state.playMode = 'computer';
      const initial = state.units.map(unit => [unit.x, unit.y, unit.hp]);
      for (let day = 0; !state.winner && day < 12; day++) {
        const turn = state.turn;
        assert.ok(game.simulateDay(state));
        assert.ok(state.winner || state.turn === turn + 1);
        const alive = state.units.filter(unit => unit.hp > 0 && !unit.exited);
        assert.equal(new Set(alive.map(unit => game.key(unit.x, unit.y))).size, alive.length);
        assert.ok(alive.every(unit => state.map[game.key(unit.x, unit.y)]));
        assert.deepEqual(game.deserialize(JSON.stringify(state)), state);
        state = game.deserialize(JSON.stringify(state));
        assert.equal(state.playMode, 'computer');
      }
      assert.ok(state.winner);
      assert.notDeepEqual(state.units.slice(0, 14).map(unit => [unit.x, unit.y, unit.hp]), initial);
      const finished = JSON.stringify(state);
      assert.equal(game.simulateDay(state), false);
      assert.equal(JSON.stringify(state), finished);
    }
  }
});

test('spectator mode rejects malformed saves and simulation on human campaigns', () => {
  for (const ruleset of ['classic', 'manual']) {
    const state = game.createGame(1, ruleset);
    assert.throws(() => game.simulateDay(state), /Computer vs Computer/);
    state.playMode = 'invalid';
    assert.throws(() => game.deserialize(JSON.stringify(state)), /valid/i);
    state.playMode = 'computer';
    state.phase = 'american';
    if (ruleset === 'manual') state.playerSide = 'american';
    assert.throws(() => game.deserialize(JSON.stringify(state)), /valid/i);
  }
});

test('campaign scoreboard reports threshold standings, uncapped points and actual victory in both modes', () => {
  for (const ruleset of ['classic', 'manual']) {
    const state = game.createGame(1, ruleset);
    for (const points of [0, 7, 8, 15]) {
      state.score = points;
      const standing = game.campaignStanding(state);
      assert.equal(standing.points, points);
      assert.equal(standing.target, 8);
      assert.equal(standing.leader, points >= 8 ? 'german' : 'american');
      assert.equal(standing.provisional, true);
    }
    state.score = 15; state.winner = 'american';
    const final = game.campaignStanding(state);
    assert.equal(final.status, 'AMERICAN VICTORY');
    assert.equal(final.provisional, false);
  }
});

test('scenario has a valid hex map, unique counters and connected initial supply', () => {
  const state = game.createGame();
  assert.equal(Object.keys(state.map).length, 180);
  assert.equal(new Set(state.units.map(unit => game.key(unit.x, unit.y))).size, state.units.length);
  assert.ok(state.units.filter(unit => unit.side === 'german').every(unit => unit.supplied));
  const rivers = Object.values(state.map).filter(tile => tile.terrain === 'river');
  assert.ok(rivers.some(tile => tile.riverSize === 'major'));
  assert.ok(rivers.some(tile => tile.riverSize === 'minor'));
  assert.ok(rivers.every(tile => tile.riverSize === (tile.col === 3 ? 'major' : 'minor')));
  assert.deepEqual(game.riverCourses.map(course => [course.size, course.points.length]),
    [['major', 38], ['minor', 23], ['minor', 38]]);
  assert.ok(game.riverCourses.every(course => course.points.slice(1).every(([x, y], index) => {
    const [previousX, previousY] = course.points[index];
    return Math.abs(Math.hypot(x - previousX, y - previousY) - 26) < 2;
  })), 'River lines must follow connected hex edges');
  assert.equal(game.neighbors(state, state.map['14,6']).length, 6);
});
test('movement respects terrain budget, occupied tiles and the active side', () => {
  const state = game.createGame();
  const unit = state.units[0];
  const [targetKey, option] = [...game.reachable(state, unit)][0];
  const target = state.map[targetKey];
  assert.ok(game.move(state, unit.id, target.x, target.y));
  assert.ok(unit.mp <= option.remaining);
  assert.equal(game.move(state, unit.id, -1, -1), false);
  assert.equal(game.reachable(state, state.units.find(counter => counter.side === 'american')).size, 0);
  assert.equal(game.reachable(state, unit).has(game.key(state.units[1].x, state.units[1].y)), false);
});
test('combat is adjacent, deterministic and limited to one attack', () => {
  const state = game.createGame(123);
  const attacker = state.units[0], defender = state.units.find(unit => unit.id === 'a1');
  defender.x = attacker.x + 2; defender.y = attacker.y;
  const copy = game.deserialize(JSON.stringify(state));
  assert.ok(game.attack(state, attacker.id, defender.id));
  assert.equal(game.attack(state, attacker.id, defender.id), false);
  assert.equal(attacker.mp, 0);
  assert.ok(game.attack(copy, attacker.id, defender.id));
  assert.deepEqual(copy.units, state.units);
  assert.equal(game.attack(state, 'g2', 'a7'), false);
});
test('enemy ring cuts supply and snow reduces movement', () => {
  const state = game.createGame();
  const unit = state.units[0]; unit.x = 14; unit.y = 6;
  const ring = game.neighbors(state, unit);
  const enemies = state.units.filter(counter => counter.side === 'american');
  ring.forEach((tile, index) => { enemies[index].x = tile.x; enemies[index].y = tile.y; });
  assert.equal(game.isSupplied(state, unit), false);
  state.weather = 'Snow'; game.beginSide(state, 'german');
  assert.equal(unit.mp, unit.mobility - 3);
});
test('full campaign advances, adds reinforcements and ends on turn twelve', () => {
  const state = game.createGame();
  for (let count = 0; count < 12 && !state.winner; count++) {
    assert.ok(game.endTurn(state));
    const living = state.units.filter(unit => unit.hp > 0);
    assert.equal(new Set(living.map(unit => game.key(unit.x, unit.y))).size, living.length);
    assert.ok(living.every(unit => state.map[game.key(unit.x, unit.y)]));
    assert.doesNotThrow(() => game.deserialize(JSON.stringify(state)));
  }
  assert.equal(state.turn, 12);
  assert.equal(state.winner, 'american');
  assert.equal(state.weather, 'Clear');
  assert.ok(state.units.some(unit => unit.id === 'a10'));
  assert.equal(game.endTurn(state), false);
});
test('American command can hand off turns to the German AI', () => {
  const state = game.createGame();
  game.computerTurn(state, 'german');
  assert.equal(state.phase, 'german');
  game.beginSide(state, 'american');
  assert.equal(state.phase, 'american');
  assert.ok(state.units.filter(unit => unit.side === 'american' && unit.hp > 0).every(unit => game.reachable(state, unit).size > 0));
  assert.ok(game.endTurn(state, 'american'));
  assert.equal(state.phase, 'american');
  assert.equal(state.turn, 2);
});
test('save import round-trips and rejects malformed data', () => {
  const state = game.createGame();
  assert.deepEqual(game.deserialize(JSON.stringify(state)), state);
  assert.throws(() => game.deserialize('{broken'));
  const invalid = structuredClone(state); invalid.units[0].x = 1000;
  assert.throws(() => game.deserialize(JSON.stringify(invalid)), /unit data/);
  invalid.units[0].x = state.units[0].x; invalid.units[0].hp = -1;
  assert.throws(() => game.deserialize(JSON.stringify(invalid)), /unit data/);
});

test('manual mode preserves classic defaults and starts with German combat, not movement', () => {
  assert.equal(game.createGame().version, 1);
  assert.throws(() => game.createGame(1, 'unknown'), /ruleset/);
  const state = game.createGame(1, 'manual', 'german');
  assert.equal(state.version, 2);
  assert.equal(state.orderPhase, 'combat');
  assert.equal(state.phase, 'german');
  assert.ok(state.units.filter(unit => unit.side === 'german').every(unit => unit.mp === 0));
  assert.ok(state.units.every(unit => unit.mobility === (unit.type === 'armor' ? 10 : 8)));
  assert.equal(game.reachable(state, state.units[0]).size, 0);
  assert.ok(game.manual.eligibleAttackers(state, state.units.find(unit => unit.id === 'a1')).length);
  assert.deepEqual(game.deserialize(JSON.stringify(state)), state);
  const american = game.createGame(1, 'manual', 'american');
  assert.equal(american.phase, 'american');
  assert.equal(american.orderPhase, 'movement');
  assert.ok(american.units.filter(unit => unit.side === 'american').every(unit => unit.mp <= unit.mobility / 2));
});

test('manual turn-two movement budget matches five armor or four infantry clear hexes', () => {
  const state = game.createGame(1, 'manual');
  state.turn = 2;
  state.weather = 'Overcast';
  game.manual.beginSide(state, 'german');
  assert.ok(state.units.filter(unit => unit.side === 'german' && unit.type === 'armor')
    .every(unit => unit.mp === 10));
  assert.ok(state.units.filter(unit => unit.side === 'german' && unit.type === 'infantry')
    .every(unit => unit.mp === 8));
});

test('manual reinforcement schedule exposes reference-card dates, strengths and arrival status', () => {
  const state = game.createGame(1, 'manual');
  const schedule = game.manual.reinforcementSchedule(state);
  assert.deepEqual(schedule.map(({ date, name, strength, status }) => [date, name, strength, status]), [
    ['16 DEC', '101 Airborne', 70, 'On map'],
    ['16 DEC', '82 Airborne', 70, 'On map'],
    ['16 DEC', '116 Panzer', 100, 'On map'],
    ['18 DEC', '2 SS Panzer', 115, 'Scheduled'],
    ['18 DEC', '9 SS Panzer', 115, 'Scheduled'],
    ['20 DEC', '3 Armored', 100, 'Scheduled'],
    ['20 DEC', '4 Armored', 100, 'Scheduled']
  ]);
  state.units.find(unit => unit.id === 'a8').hp = 0;
  assert.equal(game.manual.reinforcementSchedule(state).find(entry => entry.id === 'a8').status, 'Eliminated');
  assert.deepEqual(game.manual.offMapReinforcements(state).map(entry => entry.id),
    ['g10', 'g9', 'a11', 'a10']);
  state.turn = 3;
  assert.equal(game.manual.reinforcementSchedule(state).find(entry => entry.id === 'g9').status, 'Due after combat');
  state.pending.push('g9');
  assert.equal(game.manual.reinforcementSchedule(state).find(entry => entry.id === 'g9').status, 'Delayed');
  assert.equal(game.manual.offMapReinforcements(state).find(entry => entry.id === 'g9').status, 'Delayed');
});

test('manual mode advances movement to combat without advancing the day for either side', () => {
  for (const side of ['german', 'american']) {
    const state = game.createGame(1, 'manual', side);
    if (state.orderPhase === 'combat') game.endTurn(state, side);
    const day = state.turn;
    assert.equal(state.orderPhase, 'movement');
    assert.equal(game.attack(state, side === 'german' ? 'g1' : 'a1', side === 'german' ? 'a1' : 'g1'), false);
    assert.ok(game.endTurn(state, side));
    assert.equal(state.turn, day);
    assert.equal(state.orderPhase, 'combat');
    assert.equal(game.reachable(state, state.units.find(unit => unit.side === side)).size, 0);
    assert.ok(game.endTurn(state, side));
    assert.equal(state.turn, day + 1);
    assert.equal(state.phase, side);
    assert.equal(state.orderPhase, 'movement');
  }
});

test('manual local supply uses consecutive gaps, friendly support and airdrops', () => {
  const state = game.createGame(1, 'manual', 'german', 1);
  const unit = state.units.find(unit => unit.id === 'g1');
  unit.x = 14; unit.y = 6;
  const ring = [[13, 5], [15, 5], [16, 6], [15, 7], [13, 7], [12, 6]];
  const enemies = state.units.filter(unit => unit.side === 'american');
  [0, 2, 4].forEach((index, count) => { enemies[count].x = ring[index][0]; enemies[count].y = ring[index][1]; });
  assert.equal(game.manual.localSupply(state, unit), 'isolated');
  const friends = state.units.filter(candidate => candidate.side === 'german' && candidate.id !== unit.id);
  friends[0].x = ring[1][0]; friends[0].y = ring[1][1];
  assert.equal(game.manual.localSupply(state, unit), 'limited');
  friends[1].x = ring[3][0]; friends[1].y = ring[3][1];
  assert.equal(game.manual.localSupply(state, unit), 'supplied');
  friends[0].hp = 0; friends[1].hp = 0;
  assert.equal(game.manual.localSupply(state, unit), 'isolated');
  assert.ok(!game.manual.eligibleAttackers(state, enemies[0]).includes(unit));
  unit.airdropTurn = state.turn;
  assert.equal(game.manual.localSupply(state, unit), 'isolated', 'Airdrops cannot supply German units');
  state.units.forEach(counter => { counter.side = counter.side === 'german' ? 'american' : 'german'; });
  assert.equal(game.manual.localSupply(state, unit), 'limited');
  delete unit.airdropTurn;
});

test('manual movement consumes terrain points, fuel and readiness while replacements are capped', () => {
  const state = game.createGame(1, 'manual');
  state.turn = 2;
  game.beginSide(state, 'german');
  const unit = state.units.find(unit => unit.id === 'g4');
  const option = [...game.reachable(state, unit)].find(([, entry]) => entry.path.length === 1);
  assert.ok(option);
  const [key, entry] = option, tile = state.map[key], readiness = unit.readiness, fuel = state.fuel;
  assert.ok(game.move(state, unit.id, tile.x, tile.y));
  assert.equal(unit.readiness, Math.max(0, readiness - 2));
  assert.equal(state.fuel, fuel - Math.ceil(game.manual.COST[tile.terrain] / 2));
  assert.equal(unit.mp, entry.remaining);
  assert.ok(state.units.every(unit => unit.hp <= unit.maxHp));
});

test('manual zone entry costs four extra MP and armor alone can continue through a zone', () => {
  const state = game.createGame(1, 'manual', 'german', 1);
  state.units.forEach(unit => { unit.hp = 0; });
  Object.values(state.map).forEach(tile => { tile.terrain = 'river'; });
  const unit = state.units.find(unit => unit.id === 'g1'), enemy = state.units.find(unit => unit.id === 'a1');
  Object.assign(unit, { hp: 100, x: 14, y: 6, mp: 16 });
  Object.assign(enemy, { hp: 70, x: 17, y: 5 });
  state.orderPhase = 'movement';
  state.map['16,6'].terrain = 'plain';
  state.map['18,6'].terrain = 'plain';
  assert.equal(game.reachable(state, unit).get('16,6').remaining, 10, 'Plain costs 2 plus 4 for ZOC entry');
  assert.equal(game.reachable(state, unit).get('18,6').remaining, 0, 'Next plain costs 2 plus 8 between ZOCs');
  unit.type = 'infantry';
  assert.ok(!game.reachable(state, unit).has('18,6'));
  assert.ok(game.move(state, unit.id, 16, 6));
  assert.equal(unit.mp, 0);
});

test('manual general supply and fuel never mutate when an invalid order is rejected', () => {
  const state = game.createGame(1, 'manual');
  state.orderPhase = 'movement';
  const before = JSON.stringify(state);
  assert.equal(game.move(state, 'g1', -1, -1), false);
  assert.equal(game.attack(state, 'g1', 'a1'), false);
  assert.equal(JSON.stringify(state), before);
});

test('manual combined combat validates artillery, preserves seeded results and reports all losses', () => {
  const state = game.createGame(123, 'manual');
  const defender = state.units.find(unit => unit.id === 'a1');
  const extra = state.units.find(unit => unit.id === 'g7');
  extra.x = defender.x + 1; extra.y = defender.y + 1;
  const attackers = game.manual.eligibleAttackers(state, defender).map(unit => unit.id);
  assert.ok(attackers.length >= 2);
  const before = JSON.stringify(state);
  assert.throws(() => game.manual.resolveAttack(state, attackers, defender.id, { artillery: 7 }), /artillery/);
  assert.equal(JSON.stringify(state), before, 'Invalid combat must not mutate the state');
  assert.throws(() => game.manual.resolveAttack(state, [attackers[0], attackers[0]], defender.id), /eligible/);
  const copy = game.deserialize(before);
  state.generalSupply.german = 0; copy.generalSupply.german = 0;
  const options = { strategy: 'major', artillery: 2 };
  const result = game.manual.resolveAttack(state, attackers, defender.id, options);
  assert.deepEqual(game.manual.resolveAttack(copy, attackers, defender.id, options), result);
  assert.deepEqual(copy, state);
  assert.equal(state.artillery.german, 4);
  assert.equal(state.artillery.american, 2);
  assert.equal(result.units.length, attackers.length + 1);
  assert.ok(result.events.some(event => event.includes('general supply failed')));
  assert.ok(result.units.every(unit => unit.hp - unit.remaining === unit.loss));
  assert.throws(() => game.manual.resolveAttack(state, attackers, defender.id), /eligible/);
  assert.deepEqual(game.deserialize(JSON.stringify(state)), state);
});

test('legacy manual blocked reinforcement waits one phase before alternative entry and both sides arrive', () => {
  const state = game.createGame(1, 'manual', 'german', 1);
  // Remove the already-arrived day-one unit to exercise blocked entry independently.
  state.units = state.units.filter(unit => unit.id !== 'a8');
  state.arrivals = state.arrivals.filter(id => id !== 'a8');
  state.turn = 3;
  const blocker = state.units.find(unit => unit.id === 'a7');
  blocker.x = 10; blocker.y = 8;
  game.manual.reinforce(state);
  assert.ok(!state.units.some(unit => unit.id === 'a8'));
  assert.ok(state.pending.includes('a8'));
  assert.ok(state.units.some(unit => unit.id === 'g8'));
  state.turn = 4;
  game.manual.reinforce(state);
  const airborne = state.units.find(unit => unit.id === 'a8');
  assert.ok(airborne);
  assert.notEqual(game.key(airborne.x, airborne.y), game.key(blocker.x, blocker.y));
  assert.deepEqual(game.arrivalsThisTurn(state).map(unit => unit.id), ['a8']);
  assert.deepEqual(game.arrivalsThisTurn(game.deserialize(JSON.stringify(state))).map(unit => unit.id), ['a8']);
  state.turn = 5; game.manual.reinforce(state);
  assert.ok(state.units.some(unit => unit.id === 'a9'));
  assert.ok(state.units.some(unit => unit.id === 'g9'));
});

test('campaign seeds validate strictly and reproduce complete computer campaigns', () => {
  for (const value of [-1, 1.5, NaN, Infinity, 4294967296, 'abc', '1e3', '']) {
    assert.throws(() => game.createGame(value), /Seed/);
  }
  assert.equal(game.parseSeed('0'), 0);
  assert.equal(game.parseSeed('4294967295'), 4294967295);
  for (const ruleset of ['classic', 'manual']) {
    const first = game.createGame(42, ruleset), second = game.createGame(42, ruleset);
    first.playMode = second.playMode = 'computer';
    while (!first.winner) { game.simulateDay(first); game.simulateDay(second); }
    assert.deepEqual(second, first);
    assert.equal(first.initialSeed, 42);
    assert.equal(game.deserialize(JSON.stringify(first)).initialSeed, 42);
    const different = game.createGame(43, ruleset); different.playMode = 'computer';
    while (!different.winner) game.simulateDay(different);
    assert.notDeepEqual(different.units, first.units);
  }
});

test('spectator actions resume after every order and equal whole-day simulation', () => {
  for (const ruleset of ['classic', 'manual']) {
    for (const seed of [1, 42, 123]) {
      let stepped = game.createGame(seed, ruleset), daily = structuredClone(stepped);
      stepped.playMode = daily.playMode = 'computer';
      const actions = { movement: 0, battle: 0, american: 0, days: 0 };
      for (let count = 0; !stepped.winner; count++) {
        assert.ok(count < 1000, 'Action stepping must terminate');
        const action = game.simulateAction(stepped);
        assert.ok(action && action.text);
        if (action.path && !action.battle) actions.movement++;
        if (action.battle) actions.battle++;
        if (action.side === 'american') actions.american++;
        if (action.dayEnd) {
          actions.days++; game.simulateDay(daily);
          assert.deepEqual(stepped, daily, 'Stepping and Next Day must use the same orders');
        }
        const restored = game.deserialize(JSON.stringify(stepped));
        assert.deepEqual(restored, stepped);
        stepped = restored;
        const active = stepped.units.filter(unit => unit.hp > 0 && !unit.exited);
        assert.equal(new Set(active.map(unit => game.key(unit.x, unit.y))).size, active.length);
      }
      assert.ok(actions.movement > 0 && actions.battle > 0 && actions.american > 0 && actions.days > 0);
      assert.equal(game.simulateAction(stepped), false);
    }
  }
});

test('playback and report metadata reject invalid or incompatible imports', () => {
  for (const ruleset of ['classic', 'manual']) {
    const original = game.createGame(42, ruleset); original.playMode = 'computer';
    game.simulateAction(original);
    for (const patch of [
      { initialSeed: '42' }, { initialSeed: -1 },
      { computerProgress: { stage: 'invalid', index: 0 } },
      { computerProgress: { stage: 'combat', index: 99 } },
      { reports: { german: [{ text: 'future event', turn: 99 }], american: [] } },
      { daySummary: { turn: 99, german: '', american: '' } },
      { dayStart: { german: { points: -1 }, american: {} } }
    ]) {
      assert.throws(() => game.deserialize(JSON.stringify({ ...original, ...patch })), /invalid|Seed/i);
    }
    const human = { ...original, playMode: 'human' };
    assert.throws(() => game.deserialize(JSON.stringify(human)), /invalid/i);
  }
});

test('battle previews explain modifiers without spending resources or randomness', () => {
  for (const ruleset of ['classic', 'manual']) {
    const state = game.createGame(42, ruleset);
    if (ruleset === 'classic') {
      const attacker = state.units.find(unit => unit.id === 'g1');
      const defender = state.units.find(unit => unit.id === 'a1');
      attacker.x = defender.x + 2; attacker.y = defender.y;
    }
    const defender = state.units.find(unit => unit.id === 'a1');
    const ids = ruleset === 'manual' ? game.manual.eligibleAttackers(state, defender).map(unit => unit.id) : ['g1'];
    const before = JSON.stringify(state);
    const preview = game.battlePreview(state, ids, defender.id, { strategy: 'major', artillery: 1 });
    assert.ok(preview.odds > 0);
    assert.match(preview.detail, /supply/i);
    assert.match(preview.detail, /defense/i);
    assert.match(preview.risk, /loss|uncertain|risk|guaranteed/i);
    assert.equal(JSON.stringify(state), before);
    assert.throws(() => game.battlePreview(state, [], defender.id), /eligible/);
    assert.throws(() => game.battlePreview(state, ['missing'], defender.id), /eligible/);
    if (ruleset === 'manual') {
      const none = game.battlePreview(state, ids, defender.id, { artillery: 0 });
      assert.ok(preview.odds > none.odds);
      assert.throws(() => game.battlePreview(state, ids, defender.id, { artillery: 99 }), /artillery/);
    }
    assert.equal(JSON.stringify(state), before);
  }
});

test('order guidance identifies phase, spent movement, commitment and isolation', () => {
  const state = game.createGame(1, 'manual'), unit = state.units[0];
  assert.match(game.orderReason(state, unit, 'move'), /no movement on day 1/i);
  state.turn = 2;
  assert.match(game.orderReason(state, unit, 'move'), /Movement is closed/);
  state.orderPhase = 'combat';
  unit.fired = true;
  assert.match(game.orderReason(state, unit, 'attack'), /already attacked/);
  unit.fired = false; state.orderPhase = 'movement'; unit.mp = 0;
  assert.match(game.orderReason(state, unit, 'move'), /no movement points remaining/i);
  assert.match(game.orderReason(state, unit, 'attack'), /End Movement/);
  state.playMode = 'computer';
  assert.match(game.orderReason(state, unit, 'move'), /Spectator/);
});

test('new manual campaigns use card data while legacy manual saves retain old terrain and dates', () => {
  const state = game.createGame(1, 'manual');
  const expected = { g1: 115, g2: 115, g3: 100, g4: 100, g5: 65, g6: 50, g7: 55,
    a1: 75, a2: 55, a3: 70, a4: 120, a5: 115, a6: 75, a7: 60, a8: 70, a9: 70, g8: 100 };
  for (const [id, hp] of Object.entries(expected)) {
    assert.equal(state.units.find(unit => unit.id === id).maxHp, hp);
  }
  assert.deepEqual(game.manual.COST, { plain: 2, road: 1, town: 3, forest: 6, ridge: 4, river: 8, westwall: 4 });
  for (const [terrain, defense] of [['forest', 2], ['ridge', 1.5], ['town', 2.25], ['plain', 1]]) {
    assert.equal(game.manual.terrainDefense(state, { terrain }), defense);
  }
  assert.ok(game.arrivalsThisTurn(state).some(unit => unit.id === 'a8'));
  state.turn = 3; game.manual.reinforce(state);
  assert.equal(state.units.find(unit => unit.id === 'g9').maxHp, 115);
  assert.equal(state.units.find(unit => unit.id === 'g10').maxHp, 115);
  assert.ok(!state.units.some(unit => unit.id === 'a10'));
  state.turn = 5; game.manual.reinforce(state);
  assert.ok(state.units.some(unit => unit.id === 'a10') && state.units.some(unit => unit.id === 'a11'));
  assert.equal(state.units.find(unit => unit.id === 'a10').mobility, 10);
  assert.equal(state.units.find(unit => unit.id === 'a11').mobility, 10);

  const legacy = game.createGame(1, 'manual', 'german', 1);
  legacy.units.forEach(unit => { unit.mobility = unit.type === 'armor' ? 24 : 20; });
  delete legacy.referenceVersion; delete legacy.initialSeed;
  delete legacy.dayStart; delete legacy.reports;
  legacy.units = legacy.units.filter(unit => !['a8', 'a9', 'g8'].includes(unit.id));
  legacy.arrivals = [];
  const restored = game.deserialize(JSON.stringify(legacy));
  assert.deepEqual(restored, legacy);
  assert.equal(game.manual.terrainDefense(restored, { terrain: 'town' }), 1.8);
  game.manual.reinforce(restored);
  assert.ok(!restored.units.some(unit => unit.id === 'g8'));
  restored.turn = 3; game.manual.reinforce(restored);
  assert.equal(restored.units.find(unit => unit.id === 'g8').maxHp, 100);
  assert.equal(restored.units.find(unit => unit.id === 'a8').maxHp, 100);
  assert.ok(!restored.units.some(unit => unit.id === 'g10'));
});

test('manual source strategy names are accepted and report actual commitment choices', () => {
  for (const strategy of ['recon', 'probe', 'assault', 'major']) {
    for (const defenseStrategy of ['counter', 'hold', 'delay', 'withdrawal']) {
      const state = game.createGame(123, 'manual'), defender = state.units.find(unit => unit.id === 'a1');
      const ids = game.manual.eligibleAttackers(state, defender).map(unit => unit.id);
      assert.ok(game.battlePreview(state, ids, defender.id, { strategy, defenseStrategy }).odds > 0);
      const result = game.manual.resolveAttack(state, ids, defender.id, { strategy, defenseStrategy });
      assert.ok(result.units.every(unit => unit.remaining >= 0 && unit.loss >= 0));
      assert.deepEqual(game.deserialize(JSON.stringify(state)), state);
    }
  }
});

test('day summaries and side narratives persist and explicitly label net changes', () => {
  for (const ruleset of ['classic', 'manual']) {
    const state = game.createGame(42, ruleset); state.playMode = 'computer';
    game.simulateDay(state);
    assert.equal(state.daySummary.turn, 1);
    for (const side of ['american', 'german']) {
      assert.match(state.daySummary[side], /net .*including arrivals/);
      assert.ok(state.reports[side].some(event => event.text === state.daySummary[side]));
      assert.ok(state.reports[side].length <= 16);
    }
    assert.deepEqual(game.deserialize(JSON.stringify(state)), state);
  }
});

test('objective-focused AI has varied outcomes across 60 seeds without changing victory targets', t => {
  for (const ruleset of ['classic', 'manual']) {
    const wins = { german: 0, american: 0 };
    for (let seed = 1; seed <= 60; seed++) {
      const state = game.createGame(seed, ruleset, 'german', 1); state.playMode = 'computer';
      while (!state.winner) game.simulateDay(state);
      wins[state.winner]++;
      assert.equal(game.campaignStanding(state).target, 8);
    }
    assert.ok(wins.german > 0 && wins.american > 0, 'Neither army should win every seeded benchmark');
    t.diagnostic(`${ruleset} legacy-map benchmark: ${JSON.stringify(wins)}; seeds 1–60; not a claim of 50/50 historical balance`);
  }
});

test('manual exits remove counters from occupation and award adapted victory points', () => {
  const state = game.createGame(1, 'manual');
  const unit = state.units.find(unit => unit.id === 'g1');
  const exit = Object.values(state.map).find(tile => tile.exit);
  unit.x = exit.x; unit.y = exit.y;
  const hp = unit.hp;
  game.manual.finishDay(state);
  assert.equal(unit.exited, true);
  assert.equal(game.occupant(state, exit), undefined);
  assert.equal(state.exitPoints, Math.ceil(hp / 10));
  assert.equal(game.reachable(state, unit).size, 0);
});

test('manual eliminated German forces cannot win on held objectives alone', () => {
  const state = game.createGame(1, 'manual');
  state.units.filter(unit => unit.side === 'german').forEach(unit => { unit.hp = 0; });
  Object.values(state.map).filter(tile => tile.value).forEach(tile => { tile.owner = 'german'; });
  game.manual.finishDay(state);
  assert.equal(state.winner, 'american');
});

test('manual saves reject malformed phase, resource, readiness and reinforcement values', () => {
  const state = game.createGame(1, 'manual');
  for (const corrupt of [
    copy => { copy.phase = 'american'; },
    copy => { copy.generalSupply.german = -1; },
    copy => { copy.fuel = '120'; },
    copy => { copy.units[0].readiness = 8; },
    copy => { copy.units[0].exited = 'false'; },
    copy => { copy.units[0] = null; },
    copy => { copy.pending = ['unknown']; },
    copy => { copy.arrivals = ['g1']; }
  ]) {
    const copy = structuredClone(state); corrupt(copy);
    assert.throws(() => game.deserialize(JSON.stringify(copy)), /Invalid/);
  }
});

test('manual campaigns preserve saves, resources and unique occupation through twelve days', () => {
  for (const side of ['german', 'american']) {
    for (const seed of [1, 42, 123]) {
      let state = game.createGame(seed, 'manual', side);
      for (let steps = 0; !state.winner && steps < 24; steps++) {
        const active = state.units.filter(unit => unit.hp > 0 && !unit.exited);
        assert.equal(new Set(active.map(unit => game.key(unit.x, unit.y))).size, active.length);
        assert.ok(state.units.every(unit => unit.hp <= unit.maxHp && unit.mp >= 0 && unit.mp <= unit.mobility));
        assert.ok(state.fuel >= 0 && state.generalSupply.german >= 0 && state.generalSupply.american >= 0);
        assert.deepEqual(game.deserialize(JSON.stringify(state)), state);
        assert.ok(game.endTurn(state, side));
        state = game.deserialize(JSON.stringify(state));
      }
      assert.ok(state.winner);
      assert.equal(state.phase, 'finished');
      assert.equal(game.endTurn(state, side), false);
    }
  }
});

test('active manual campaigns preserve invariants after combined attacks, retreats and advances', () => {
  let orders = 0;
  for (const side of ['german', 'american']) {
    for (const seed of [2, 7, 19]) {
      const state = game.createGame(seed, 'manual', side);
      const verify = () => {
        const active = state.units.filter(unit => unit.hp > 0 && !unit.exited);
        assert.equal(new Set(active.map(unit => game.key(unit.x, unit.y))).size, active.length);
        assert.ok(state.units.every(unit => state.map[game.key(unit.x, unit.y)] && unit.hp <= unit.maxHp));
        assert.deepEqual(game.deserialize(JSON.stringify(state)), state);
      };
      for (let step = 0; !state.winner && step < 24; step++) {
        if (state.orderPhase === 'movement') {
          for (const unit of state.units.filter(unit => unit.side === side && unit.hp > 0 && !unit.exited)) {
            const targets = Object.values(state.map).filter(tile => tile.value && tile.owner !== side);
            if (!targets.length) continue;
            const rank = tile => Math.min(...targets.map(target => game.distance(tile, target)));
            const tile = [...game.reachable(state, unit).keys()].map(key => state.map[key])
              .sort((a, b) => rank(a) - rank(b))[0];
            if (tile && rank(tile) < rank(unit)) {
              assert.ok(game.move(state, unit.id, tile.x, tile.y));
              orders++; verify();
            }
          }
        } else {
          for (const defender of state.units.filter(unit => unit.side !== side && unit.hp > 0 && !unit.exited)) {
            const attackers = game.manual.eligibleAttackers(state, defender);
            if (!attackers.length) continue;
            game.manual.resolveAttack(state, attackers.map(unit => unit.id), defender.id, {
              strategy: ['probe', 'assault', 'major'][step % 3],
              artillery: Math.min(attackers.length, state.artillery[side])
            });
            orders++; verify();
          }
        }
        assert.ok(game.endTurn(state, side));
        verify();
      }
      assert.ok(state.winner);
    }
  }
  assert.ok(orders > 30);
});
test('arrivals are visible on their scheduled turn and survive save restore for either side', () => {
  for (const side of ['german', 'american']) {
    const state = game.createGame();
    game.beginSide(state, side);
    const expected = { 3: ['a8'], 5: ['a9'], 7: ['a10', 'a11'] };
    for (let turn = 1; turn <= 8; turn++) {
      const arrivals = game.arrivalsThisTurn(state);
      assert.deepEqual(arrivals.map(unit => unit.id), expected[turn] || []);
      for (const unit of arrivals) {
        const tile = state.map[game.key(unit.x, unit.y)];
        assert.ok(tile);
        assert.equal(game.occupant(state, tile), unit);
      }
      const restored = game.deserialize(JSON.stringify(state));
      assert.deepEqual(game.arrivalsThisTurn(restored), arrivals);
      assert.equal(restored.phase, side);
      if (turn < 8) assert.ok(game.endTurn(state, side));
    }
    assert.ok(state.units.some(unit => unit.id === 'a8' && unit.hp > 0));
  }
});

function verifyCampaign(state) {
  const living = state.units.filter(unit => unit.hp > 0);
  assert.equal(new Set(living.map(unit => game.key(unit.x, unit.y))).size, living.length, 'Living units must not stack');
  for (const unit of state.units) {
    assert.ok(state.map[game.key(unit.x, unit.y)], `${unit.id} must remain on the map`);
    assert.ok(unit.hp >= 0 && unit.hp <= unit.maxHp, `${unit.id} strength must stay in range`);
    assert.ok(unit.mp >= 0 && unit.mp <= unit.mobility, `${unit.id} movement must stay in range`);
    if (unit.hp > 0) {
      const tile = state.map[game.key(unit.x, unit.y)];
      if (tile.value) assert.equal(tile.owner, unit.side, `${unit.id} must control its occupied objective`);
    }
  }
  const expectedScore = Object.values(state.map).reduce((total, tile) => total + (tile.owner === 'german' ? tile.value || 0 : 0), 0);
  assert.equal(state.score, expectedScore, 'Score must equal controlled objective values');
  assert.deepEqual(game.deserialize(JSON.stringify(state)), state, 'Saving and restoring must preserve every state field');
}

test('active multi-seed campaigns issue orders and preserve invariants after every action', context => {
  const totals = { campaigns: 0, moves: 0, attacks: 0, captures: 0, retreats: 0, eliminated: 0, germanWins: 0, americanWins: 0 };
  for (const seed of [1, 2, 3, 7, 19, 42, 123, 999, 1944, 19441216, 4294967294, 4294967295]) {
    let state = game.createGame(seed);
    verifyCampaign(state);
    while (!state.winner) {
      for (const unit of state.units.filter(candidate => candidate.side === 'german' && candidate.hp > 0)) {
        const targets = Object.values(state.map).filter(tile => tile.value && tile.owner !== unit.side);
        const adjacent = () => state.units.filter(enemy => enemy.side === 'american' && enemy.hp > 0 && game.distance(unit, enemy) === 1)
          .sort((first, second) => game.combatOdds(state, unit, second) - game.combatOdds(state, unit, first));
        if (!adjacent().length && targets.length) {
          const options = [...game.reachable(state, unit).keys()].map(tileKey => state.map[tileKey]);
          const rank = tile => Math.min(...targets.map(target => game.distance(tile, target))) - (tile.value && tile.owner !== unit.side ? 6 : 0);
          options.sort((first, second) => rank(first) - rank(second) || first.col - second.col);
          const destination = options[0];
          if (destination && rank(destination) < rank(unit)) {
            const priorMovement = unit.mp, priorScore = state.score;
            assert.ok(game.move(state, unit.id, destination.x, destination.y));
            assert.ok(unit.mp < priorMovement, 'A move must consume movement');
            totals.moves++;
            if (state.score > priorScore) totals.captures++;
            verifyCampaign(state);
          }
        }
        const defender = adjacent()[0];
        if (defender) {
          const before = game.key(defender.x, defender.y);
          assert.ok(game.attack(state, unit.id, defender.id));
          totals.attacks++;
          if (game.key(defender.x, defender.y) !== before) totals.retreats++;
          assert.equal(game.attack(state, unit.id, defender.id), false, 'A second attack must be rejected');
          verifyCampaign(state);
        }
      }
      const turn = state.turn;
      assert.ok(game.endTurn(state));
      assert.ok(state.winner || state.turn === turn + 1, 'End Turn must advance or finish');
      verifyCampaign(state);
      state = game.deserialize(JSON.stringify(state));
    }
    assert.equal(state.phase, 'finished');
    assert.equal(game.endTurn(state), false);
    assert.equal(game.reachable(state, state.units[0]).size, 0);
    totals.campaigns++;
    totals.eliminated += state.units.filter(unit => unit.hp === 0).length;
    totals[state.winner === 'german' ? 'germanWins' : 'americanWins']++;
    context.diagnostic(`seed=${seed} turn=${state.turn} winner=${state.winner} score=${state.score}`);
  }
  assert.ok(totals.moves > 0 && totals.attacks > 0 && totals.captures > 0 && totals.retreats > 0 && totals.eliminated > 0);
  context.diagnostic(JSON.stringify(totals));
});

test('final-turn objective threshold and elimination select the correct winner', () => {
  for (const score of [7, 8]) {
    const state = game.createGame();
    state.turn = 12; state.weather = 'Clear';
    for (const unit of state.units.filter(candidate => candidate.side === 'american')) {
      unit.mobility = 1; unit.mp = 1;
      if (unit.id === 'a7') { unit.x = 27; unit.y = 11; }
    }
    for (const tile of Object.values(state.map).filter(candidate => candidate.value)) {
      tile.owner = tile.name === 'DINANT' || tile.name === (score === 8 ? 'MEUSE' : 'BASTOGNE') ? 'german' : 'american';
    }
    state.score = score;
    assert.ok(game.endTurn(state));
    assert.equal(state.score, score);
    assert.equal(state.winner, score === 8 ? 'german' : 'american');
    verifyCampaign(state);
  }
  for (const defeated of ['german', 'american']) {
    const state = game.createGame();
    state.units.filter(unit => unit.side === defeated).forEach(unit => { unit.hp = 0; });
    assert.ok(game.endTurn(state));
    assert.equal(state.winner, defeated === 'german' ? 'american' : 'german');
    assert.equal(state.turn, 1);
    verifyCampaign(state);
  }
});

test('automatic retreats keep British formations off the Meuse until a German unit exits', () => {
  const reference = game.referenceMap, point = reference.position;
  function battle(nationality, exit) {
    const state = game.createGame(2, 'manual', 'german', 2);
    state.units.forEach(unit => { unit.hp = 0; });
    const defender = state.units.find(unit => unit.side === 'american');
    const attacker = state.units.find(unit => unit.side === 'german');
    const friends = state.units.filter(unit => unit.side === 'american' && unit !== defender);
    defender.nationality = nationality;
    Object.assign(defender, point(1, 4), { hp: 90, maxHp: 90, readiness: 0, defended: false });
    Object.assign(attacker, point(2, 3), { hp: 80, maxHp: 80, readiness: 7, fired: false });
    [point(1, 3), point(2, 5), point(1, 5), point(0, 4)].forEach((tile, index) => {
      Object.assign(friends[index], tile, { hp: 40, maxHp: 40 });
    });
    for (const tile of Object.values(state.map)) tile.terrain = 'plain';
    state.phase = 'german'; state.orderPhase = 'combat'; state.playerSide = 'american';
    state.generalSupply = { german: 120, american: 120 };
    state.artillery = { german: 0, american: 0 };
    state.weather = 'Overcast'; state.seed = 2;
    if (exit === 'points') state.exitPoints = 1;
    if (exit === 'flag') state.units.find(unit => unit.side === 'german' && unit !== attacker).exited = true;
    const result = game.manual.resolveAttack(state, [attacker.id], defender.id, {
      strategy: 'major', defenseStrategy: 'delay', artillery: 0
    });
    return { state, defender, result };
  }
  const blocked = battle('british');
  assert.equal(blocked.result.retreatPath.length, 0);
  assert.match(blocked.result.events.join(' '), /retreat 0\//);
  assert.equal(game.key(blocked.defender.x, blocked.defender.y), game.key(point(1, 4).x, 4));
  const allowed = battle('us');
  assert.equal(reference.crossesMeuse(point(1, 4), allowed.state.map[allowed.result.retreatPath[0]]), true);
  for (const exit of ['points', 'flag']) {
    const lifted = battle('british', exit);
    assert.equal(reference.crossesMeuse(point(1, 4), lifted.state.map[lifted.result.retreatPath[0]]), true);
  }
});