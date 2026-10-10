(function (root) {
  'use strict';
  const ROT = typeof module === 'object' ? require('./rot.js') : root.ROT;
  const createManualRules = typeof module === 'object' ? require('./manual-rules.js') : root.createManualRules;
  let manualRules;
  const COLS = 15, ROWS = 12, LAST_TURN = 12;
  const DIRECTIONS = ROT.DIRS[6];
  const TERRAIN = {
    plain: { cost: 1, defense: 1, symbol: '.' },
    forest: { cost: 2, defense: 1.5, symbol: '*' },
    ridge: { cost: 3, defense: 1.7, symbol: '^' },
    river: { cost: 3, defense: 1.3, symbol: '~' },
    road: { cost: 1, defense: 1, symbol: ':' },
    town: { cost: 1, defense: 1.8, symbol: '#' }
  };
  const opponent = side => side === 'german' ? 'american' : 'german';
  const REINFORCEMENTS = {
    3: [['a8', '101 Airborne', 'infantry', 5, 8, 10]],
    5: [['a9', '82 Airborne', 'infantry', 4, 3, 10]],
    7: [['a10', '4 Armored', 'armor', 2, 10, 10], ['a11', '3 Armored', 'armor', 1, 2, 10]]
  };
  const key = (x, y) => `${x},${y}`;
  const position = (col, row) => ({ x: col * 2 + row % 2, y: row });
  const distance = (first, second) => Math.max(Math.abs(first.y - second.y),
    (Math.abs(first.x - second.x) + Math.abs(first.y - second.y)) / 2);
  const neighbors = (state, tile) => DIRECTIONS.map(([dx, dy]) => state.map[key(tile.x + dx, tile.y + dy)]).filter(Boolean);
  const occupant = (state, tile) => state.units.find(unit => unit.hp > 0 && !unit.exited && unit.x === tile.x && unit.y === tile.y);
  const enemiesNear = (state, tile, side) => neighbors(state, tile).some(next => {
    const unit = occupant(state, next);
    return unit && unit.side !== side;
  });
  function reportOperation(state, side, text) {
    state.reports ||= { american: [], german: [] };
    state.reports[side].unshift({ text, turn: state.turn });
    state.reports[side] = state.reports[side].slice(0, 16);
  }
  function daySnapshot(state) {
    return Object.fromEntries(['american', 'german'].map(side => {
      const units = state.units.filter(unit => unit.side === side && unit.hp > 0 && !unit.exited);
      return [side, { strength: units.reduce((sum, unit) => sum + unit.hp, 0),
        formations: units.length, points: Object.values(state.map).reduce((sum, tile) =>
          sum + (tile.owner === side ? tile.value || 0 : 0), 0) }];
    }));
  }
  function summarizeDay(state) {
    const now = daySnapshot(state), before = state.dayStart;
    const signed = value => value > 0 ? `+${value}` : String(value);
    state.daySummary = { turn: state.turn, american: '', german: '' };
    for (const side of ['american', 'german']) {
      const changes = before ? `Changes today: ${signed(now[side].points - before[side].points)} held points, ${signed(now[side].formations - before[side].formations)} active formations, ${signed(now[side].strength - before[side].strength)} net ${state.ruleset === 'manual' ? 'CP' : 'strength'} (including arrivals, replacements and exits).` :
        'Daily changes unavailable: this older save has no starting-day totals.';
      const text = `End of ${15 + state.turn} December: ${now[side].points} held-objective points, ${now[side].formations} active formations, ${now[side].strength} ${state.ruleset === 'manual' ? 'CP' : 'strength'}. ${changes}`;
      state.daySummary[side] = text;
      reportOperation(state, side, text);
    }
    delete state.dayStart;
  }
  function parseSeed(value) {
    if (!/^\d{1,10}$/.test(String(value)) || Number(value) > 4294967295) {
      throw new Error('Seed must be a whole number from 0 to 4294967295.');
    }
    return Number(value);
  }
  function restoreMetadata(parsed, state) {
    if (parsed.initialSeed !== undefined) {
      if (!Number.isInteger(parsed.initialSeed)) throw new Error('Invalid original campaign seed.');
      state.initialSeed = parseSeed(parsed.initialSeed);
    }
    else delete state.initialSeed;
    if (parsed.reports !== undefined) {
      if (!parsed.reports || !['american', 'german'].every(side => Array.isArray(parsed.reports[side]) &&
        parsed.reports[side].length <= 16 && parsed.reports[side].every(event => event &&
          typeof event.text === 'string' && event.text.length <= 1000 &&
          Number.isInteger(event.turn) && event.turn >= 1 && event.turn <= parsed.turn))) {
        throw new Error('Invalid situation reports in save.');
      }
      state.reports = structuredClone(parsed.reports);
    } else delete state.reports;
    if (parsed.dayStart !== undefined) {
      if (!parsed.dayStart || !['american', 'german'].every(side => parsed.dayStart[side] &&
        ['strength', 'formations', 'points'].every(field => Number.isInteger(parsed.dayStart[side][field]) &&
          parsed.dayStart[side][field] >= 0 && parsed.dayStart[side][field] <= 10000))) {
        throw new Error('Invalid daily starting totals.');
      }
      state.dayStart = structuredClone(parsed.dayStart);
    } else delete state.dayStart;
    if (parsed.daySummary !== undefined) {
      if (!parsed.daySummary || !Number.isInteger(parsed.daySummary.turn) || parsed.daySummary.turn < 1 ||
        parsed.daySummary.turn > parsed.turn || !['american', 'german'].every(side =>
          typeof parsed.daySummary[side] === 'string' && parsed.daySummary[side].length <= 1000)) {
        throw new Error('Invalid daily summary.');
      }
      state.daySummary = structuredClone(parsed.daySummary);
    }
    if (parsed.computerProgress !== undefined) {
      const progress = parsed.computerProgress;
      if (parsed.playMode !== 'computer' || parsed.winner || !progress ||
        !['movement', 'combat'].includes(progress.stage) || !Number.isInteger(progress.index) ||
        progress.index < 0 || progress.index > parsed.units.filter(unit => unit.side === parsed.phase).length ||
        (parsed.ruleset !== 'manual' && progress.stage === 'combat' &&
          progress.index === parsed.units.filter(unit => unit.side === parsed.phase).length) ||
        !['german', 'american'].includes(parsed.phase) ||
        (parsed.ruleset === 'manual' && progress.stage !== parsed.orderPhase)) {
        throw new Error('Invalid computer playback progress.');
      }
      state.computerProgress = { stage: progress.stage, index: progress.index };
    }
    return state;
  }
  function armySummary(state, side) {
    const units = state.units.filter(unit => unit.side === side && unit.hp > 0 && !unit.exited);
    const strength = units.reduce((sum, unit) => sum + unit.hp, 0);
    const held = Object.values(state.map).filter(tile => tile.value && tile.owner === side);
    const objectives = held.map(tile => tile.name);
    const objectivePoints = held.reduce((sum, tile) => sum + tile.value, 0);
    const isolated = units.filter(unit => state.ruleset === 'manual' ? manualRules.localSupply(state, unit) === 'isolated' : !unit.supplied).length;
    return { formations: units.length, strength, isolated, objectives, objectivePoints };
  }
  function addLog(state, message) {
    state.log.unshift(message);
    state.log = state.log.slice(0, 60);
  }
  function makeUnit(id, name, side, type, col, row, strength) {
    return { id, name, side, type, ...position(col, row), hp: strength, maxHp: strength,
      attack: type === 'armor' ? 8 : 5, defense: type === 'armor' ? 6 : 5,
      mobility: type === 'armor' ? 5 : 4, mp: 0, fired: false, supplied: true };
  }
  function createGame(seed = 19441216, ruleset = 'classic', playerSide = 'german') {
    seed = parseSeed(seed);
    if (!['classic', 'manual'].includes(ruleset)) throw new Error('Unknown campaign ruleset.');
    const state = { version: 1, turn: 1, phase: 'german', weather: 'Fog', seed, initialSeed: seed,
      map: {}, units: [], log: [], score: 0, winner: null };
    for (let row = 0; row < ROWS; row++) {
      for (let col = 0; col < COLS; col++) {
        const tile = { ...position(col, row), col, row, terrain: 'plain', owner: null };
        if ((col * 7 + row * 11) % 13 < 4) tile.terrain = 'forest';
        if (row < 4 && col > 5 && col < 10) tile.terrain = 'ridge';
        if (col === 3 || (col === 8 && row > 4)) tile.terrain = 'river';
        if (row === 5 || row === 8 || (col === 6 && row > 3)) tile.terrain = 'road';
        state.map[key(tile.x, tile.y)] = tile;
      }
    }
    [[2, 5, 'DINANT', 4], [6, 8, 'BASTOGNE', 3], [7, 4, 'ST. VITH', 2],
      [5, 2, 'MALMEDY', 2], [0, 8, 'MEUSE', 4]].forEach(([col, row, name, value]) => {
      Object.assign(state.map[key(position(col, row).x, row)], { terrain: 'town', name, value, owner: 'american' });
    });
    state.units = [
      makeUnit('g1', '1 SS Panzer', 'german', 'armor', 13, 2, 10),
      makeUnit('g2', '12 SS Panzer', 'german', 'armor', 13, 4, 10),
      makeUnit('g3', '2 Panzer', 'german', 'armor', 12, 7, 10),
      makeUnit('g4', 'Panzer Lehr', 'german', 'armor', 13, 9, 9),
      makeUnit('g5', '26 Volksgrenadier', 'german', 'infantry', 12, 9, 8),
      makeUnit('g6', '62 Volksgrenadier', 'german', 'infantry', 12, 5, 8),
      makeUnit('g7', '277 Volksgrenadier', 'german', 'infantry', 12, 2, 8),
      makeUnit('a1', '99 Infantry', 'american', 'infantry', 10, 2, 7),
      makeUnit('a2', '106 Infantry', 'american', 'infantry', 10, 4, 7),
      makeUnit('a3', '28 Infantry', 'american', 'infantry', 10, 8, 7),
      makeUnit('a4', '7 Armored', 'american', 'armor', 7, 4, 8),
      makeUnit('a5', '9 Armored', 'american', 'armor', 8, 7, 8),
      makeUnit('a6', '4 Infantry', 'american', 'infantry', 9, 10, 8),
      makeUnit('a7', 'Bastogne Garrison', 'american', 'infantry', 6, 8, 6)
    ];
    addLog(state, '16 DEC 1944 / German offensive begins.');
    if (ruleset === 'manual') return manualRules.initialize(state, playerSide);
    beginSide(state, 'german');
    state.dayStart = daySnapshot(state);
    return state;
  }
  function random(state) {
    state.seed = (Math.imul(1664525, state.seed) + 1013904223) >>> 0;
    return state.seed / 4294967296;
  }
  function route(state, start, destination, passable) {
    const path = [];
    new ROT.Path.Dijkstra(destination.x, destination.y, (x, y) => {
      const tile = state.map[key(x, y)];
      return Boolean(tile && passable(tile));
    }, { topology: 6 }).compute(start.x, start.y, (x, y) => path.push(state.map[key(x, y)]));
    return path;
  }
  function isSupplied(state, unit) {
    if (state.ruleset === 'manual') return manualRules.localSupply(state, unit) === 'supplied';
    const edge = unit.side === 'german' ? COLS - 1 : 0;
    const destinations = Object.values(state.map).filter(tile => tile.col === edge);
    return destinations.some(destination => {
      const blocking = occupant(state, destination);
      if (blocking && blocking.side !== unit.side) return false;
      return route(state, unit, destination, tile => {
        if (tile.x === unit.x && tile.y === unit.y) return true;
        const other = occupant(state, tile);
        return (!other || other.side === unit.side) && (!enemiesNear(state, tile, unit.side) || Boolean(other));
      }).length > 0;
    });
  }
  function beginSide(state, side) {
    if (state.ruleset === 'manual') return manualRules.beginSide(state, side);
    state.phase = side;
    state.units.filter(unit => unit.side === side && unit.hp > 0).forEach(unit => {
      unit.supplied = isSupplied(state, unit);
      unit.mp = Math.max(1, unit.mobility - (state.weather === 'Snow' ? 1 : 0) - (unit.supplied ? 0 : 2));
      unit.fired = false;
      if (!unit.supplied && state.turn > 1) unit.hp = Math.max(1, unit.hp - 1);
    });
  }
  function movementCost(state, tile, unit) {
    return TERRAIN[tile.terrain].cost + (state.weather === 'Snow' && unit.type === 'armor' && tile.terrain !== 'road' && tile.terrain !== 'town' ? 1 : 0);
  }
  function reachable(state, unit) {
    if (state.ruleset === 'manual') return manualRules.reachable(state, unit);
    if (!unit || unit.hp <= 0 || unit.side !== state.phase || state.winner) return new Map();
    const startKey = key(unit.x, unit.y);
    const best = new Map([[startKey, { remaining: unit.mp, path: [] }]]);
    const pending = [state.map[startKey]];
    while (pending.length) {
      const current = pending.shift();
      const currentKey = key(current.x, current.y);
      const entry = best.get(currentKey);
      if (currentKey !== startKey && enemiesNear(state, current, unit.side)) continue;
      for (const tile of neighbors(state, current)) {
        if (occupant(state, tile)) continue;
        const remaining = entry.remaining - movementCost(state, tile, unit);
        const tileKey = key(tile.x, tile.y);
        if (remaining < 0 || (best.has(tileKey) && best.get(tileKey).remaining >= remaining)) continue;
        best.set(tileKey, { remaining, path: [...entry.path, tile] });
        pending.push(tile);
      }
    }
    best.delete(startKey);
    return best;
  }
  function updateScore(state) {
    state.score = Object.values(state.map).reduce((total, tile) => total + (tile.owner === 'german' ? tile.value || 0 : 0), 0);
  }
  function move(state, unitId, x, y) {
    if (state.ruleset === 'manual') return manualRules.move(state, unitId, x, y);
    const unit = state.units.find(candidate => candidate.id === unitId);
    const option = reachable(state, unit).get(key(x, y));
    if (!option) return false;
    for (const tile of option.path) {
      unit.x = tile.x; unit.y = tile.y;
      if (tile.value && tile.owner !== unit.side) {
        tile.owner = unit.side;
        addLog(state, `${unit.name} captures ${tile.name}.`);
        reportOperation(state, unit.side, `${unit.name} captured ${tile.name}, worth ${tile.value} objective points.`);
        reportOperation(state, opponent(unit.side), `${tile.name} was lost to ${unit.name}.`);
      }
    }
    unit.mp = enemiesNear(state, unit, unit.side) ? 0 : option.remaining;
    updateScore(state);
    const tile = state.map[key(x, y)];
    reportOperation(state, unit.side, `${unit.name} moved to column ${tile.col + 1}, row ${tile.row + 1}. ${unit.mp} movement points remain.`);
    return true;
  }
  function combatOdds(state, attacker, defender) {
    if (state.ruleset === 'manual') return manualRules.combatOdds(state, attacker, defender);
    const attack = attacker.attack * attacker.hp / attacker.maxHp * (attacker.supplied ? 1 : 0.5);
    const tile = state.map[key(defender.x, defender.y)];
    const defense = defender.defense * defender.hp / defender.maxHp * TERRAIN[tile.terrain].defense * (defender.supplied ? 1 : 0.75);
    return attack / Math.max(0.1, defense);
  }
  function orderReason(state, unit, order) {
    if (state.winner) return 'The campaign has finished.';
    if (state.playMode === 'computer') return 'Spectator mode: the computers issue orders; you can inspect either army.';
    if (!unit || unit.hp <= 0 || unit.exited) return 'Select a living formation.';
    if (unit.side !== state.phase) return 'Wait for this army’s orders phase.';
    if (state.ruleset === 'manual' && state.orderPhase !== (order === 'move' ? 'movement' : 'combat')) {
      return order === 'move' ? 'Movement is closed. End combat to reach the next day.' : 'End Movement before issuing attacks.';
    }
    if (unit.fired) return 'This formation has already attacked and spent its movement today.';
    if (order === 'move' && !reachable(state, unit).size) return 'No affordable empty destinations remain. Terrain, enemy zones or spent movement block this order.';
    if (order === 'attack' && state.ruleset === 'manual' && manualRules.localSupply(state, unit) === 'isolated') {
      return 'This formation is isolated and cannot attack. Open a supply gap or bring friendly support.';
    }
    return '';
  }
  function validTargets(state, unit) {
    if (!unit || unit.hp <= 0 || unit.exited || unit.fired || unit.side !== state.phase || state.winner) return [];
    return state.units.filter(enemy => enemy.hp > 0 && !enemy.exited && enemy.side !== unit.side &&
      distance(unit, enemy) === 1 && (state.ruleset !== 'manual' ||
        manualRules.eligibleAttackers(state, enemy).some(attacker => attacker.id === unit.id)));
  }
  function battlePreview(state, ids, defenderId, options = {}) {
    const defender = state.units.find(unit => unit.id === defenderId);
    const attackers = ids.map(id => state.units.find(unit => unit.id === id));
    if (!defender || !ids.length || new Set(ids).size !== ids.length ||
      attackers.some(unit => !unit || !validTargets(state, unit).includes(defender))) {
      throw new Error('Choose eligible adjacent attackers and a valid enemy target.');
    }
    if (state.ruleset === 'manual') return manualRules.preview(state, ids, defenderId, options);
    if (ids.length !== 1) throw new Error('Classic attacks use one formation.');
    const attacker = attackers[0], odds = combatOdds(state, attacker, defender);
    const tile = state.map[key(defender.x, defender.y)];
    const rolls = [0, 0.999999];
    const defenderLoss = rolls.map(roll => Math.min(defender.hp, Math.max(1, Math.min(5,
      Math.floor(odds * 1.5 + roll * 3))) + (state.weather === 'Clear' && attacker.side === 'american' ? 1 : 0)));
    const attackerLoss = rolls.map(roll => Math.min(attacker.hp, Math.max(0, Math.min(4,
      Math.floor(2 / Math.max(odds, 0.25) + (1 - roll))))));
    return { odds, detail: `${attacker.name}: ${attacker.hp}/${attacker.maxHp} strength, attack ${attacker.attack}, supply x${attacker.supplied ? 1 : 0.5}. ${defender.name}: ${defender.hp}/${defender.maxHp}, defense ${defender.defense}, ${tile.terrain} x${TERRAIN[tile.terrain].defense}, supply x${defender.supplied ? 1 : 0.75}. Weather: ${state.weather}.`,
      risk: `${odds < 1 ? 'High risk: defender has the stronger effective force.' : odds < 1.7 ? 'Contested battle: retreat is uncertain.' : 'Favorable odds, not a guaranteed victory.'} Initial losses: attacker ${Math.min(...attackerLoss)}–${Math.max(...attackerLoss)}, defender ${Math.min(...defenderLoss)}–${Math.max(...defenderLoss)}. Blocked retreats can cause 2 more losses. The attacker spends all movement; Classic has no automatic advance.` };
  }
  function attack(state, attackerId, defenderId) {
    if (state.ruleset === 'manual') {
      const defender = state.units.find(unit => unit.id === defenderId);
      if (!manualRules.eligibleAttackers(state, defender).some(unit => unit.id === attackerId)) return false;
      manualRules.resolveAttack(state, [attackerId], defenderId);
      return true;
    }
    const attacker = state.units.find(unit => unit.id === attackerId);
    const defender = state.units.find(unit => unit.id === defenderId);
    if (state.winner || !attacker || !defender || attacker.hp <= 0 || defender.hp <= 0 || attacker.side !== state.phase ||
      attacker.side === defender.side || attacker.fired || distance(attacker, defender) !== 1) return false;
    const odds = combatOdds(state, attacker, defender);
    const beforeAttack = attacker.hp, beforeDefense = defender.hp;
    const originalDefense = key(defender.x, defender.y);
    const roll = random(state);
    let defenderLoss = Math.max(1, Math.min(5, Math.floor(odds * 1.5 + roll * 3)));
    let attackerLoss = Math.max(0, Math.min(4, Math.floor(2 / Math.max(odds, 0.25) + (1 - roll))));
    if (state.weather === 'Clear' && attacker.side === 'american') defenderLoss += 1;
    attacker.hp = Math.max(0, attacker.hp - attackerLoss);
    defender.hp = Math.max(0, defender.hp - defenderLoss);
    attacker.fired = true; attacker.mp = 0;
    addLog(state, `${attacker.name} > ${defender.name}: ${odds.toFixed(1)}:1 / losses ${attackerLoss}:${defenderLoss}.`);
    if (defender.hp > 0 && attacker.hp > 0 && odds + roll > 1.7) {
      const retreat = neighbors(state, defender).filter(tile => !occupant(state, tile) && !enemiesNear(state, tile, defender.side))
        .sort((first, second) => defender.side === 'german' ? second.col - first.col : first.col - second.col)[0];
      if (retreat) { defender.x = retreat.x; defender.y = retreat.y; addLog(state, `${defender.name} retreats.`); }
      else { defender.hp = Math.max(0, defender.hp - 2); addLog(state, `${defender.name}: retreat blocked.`); }
    }
    if (attacker.hp === 0) addLog(state, `${attacker.name} eliminated.`);
    if (defender.hp === 0) addLog(state, `${defender.name} eliminated.`);
    reportOperation(state, attacker.side, `${attacker.name} attacked ${defender.name} at ${odds.toFixed(1)}:1 odds. It lost ${beforeAttack - attacker.hp} strength and inflicted ${beforeDefense - defender.hp}. ${attacker.hp === 0 ? 'Our formation was eliminated.' : `It has ${attacker.hp} strength left and cannot move or attack again today.`}`);
    reportOperation(state, defender.side, `${defender.name} was attacked by ${attacker.name}. It lost ${beforeDefense - defender.hp} strength. ${defender.hp === 0 ? 'Our formation was eliminated.' : originalDefense !== key(defender.x, defender.y) ? `It retreated and has ${defender.hp} strength left.` : `It held its position with ${defender.hp} strength left.`}`);
    return true;
  }
  function reinforcements(state) {
    for (const [id, name, type, col, row, strength] of REINFORCEMENTS[state.turn] || []) {
      const desired = state.map[key(position(col, row).x, row)];
      const tile = [desired, ...Object.values(state.map).filter(candidate => candidate.col < 5)
        .sort((first, second) => distance(first, desired) - distance(second, desired))].find(candidate => !occupant(state, candidate));
      if (tile) {
        const unit = makeUnit(id, name, 'american', type, tile.col, tile.row, strength);
        state.units.push(unit); addLog(state, `${name} arrives.`);
        reportOperation(state, unit.side, `${name} arrived as a reinforcement with ${strength} strength at column ${tile.col + 1}, row ${tile.row + 1}. Look for NEW on the map.`);
      }
    }
  }
  function arrivalsThisTurn(state) {
    if (state.ruleset === 'manual') return state.units.filter(unit => unit.hp > 0 && !unit.exited && state.arrivals.includes(unit.id));
    const ids = new Set((REINFORCEMENTS[state.turn] || []).map(([id]) => id));
    return state.units.filter(unit => unit.hp > 0 && ids.has(unit.id));
  }
  function computerTurn(state, side = 'american') {
    if (state.ruleset === 'manual') return manualRules.computerTurn(state, side);
    if (state.winner) return;
    beginSide(state, side);
    for (const unit of state.units.filter(candidate => candidate.side === side && candidate.hp > 0)) {
      const movement = planMovement(state, unit);
      if (movement) move(state, unit.id, movement.tile.x, movement.tile.y);
      const battle = planBattle(state, unit);
      if (battle) attack(state, unit.id, battle.defender.id);
    }
  }
  function planMovement(state, unit) {
    const options = reachable(state, unit);
    if (!options.size) return null;
    const enemies = state.units.filter(enemy => enemy.side !== unit.side && enemy.hp > 0 && !enemy.exited);
    const friends = state.units.filter(friend => friend.side === unit.side && friend.hp > 0 && !friend.exited && friend.id !== unit.id);
    const objectives = Object.values(state.map).filter(tile => tile.value);
    const current = state.map[key(unit.x, unit.y)], weak = unit.hp < unit.maxHp * 0.35;
    const approaching = enemies.filter(enemy => state.map[key(enemy.x, enemy.y)].col <= 8 ||
      objectives.some(tile => tile.owner === unit.side && distance(enemy, tile) <= 4));
    if (current.value && current.owner === unit.side && !weak &&
      (unit.side === 'german' || !approaching.length || enemiesNear(state, current, unit.side))) return null;
    let targets = weak ? objectives.filter(tile => tile.owner === unit.side && !enemiesNear(state, tile, unit.side)) :
      objectives.filter(tile => tile.owner !== unit.side || !occupant(state, tile));
    if (!weak && unit.side === 'american' && approaching.length) targets = approaching;
    if (!targets.length) targets = Object.values(state.map).filter(tile => tile.col === (unit.side === 'german' && !weak ? 0 : unit.side === 'german' ? 14 : 0));
    const target = targets.slice().sort((a, b) =>
      (distance(unit, a) - (a.value || 0) * 0.6) - (distance(unit, b) - (b.value || 0) * 0.6))[0];
    const rank = tile => {
      const adjacent = enemies.filter(enemy => distance(tile, enemy) === 1);
      const support = friends.filter(friend => distance(tile, friend) <= 2).length;
      const terrain = state.ruleset === 'manual' ? manualRules.terrainDefense(state, tile) : TERRAIN[tile.terrain].defense;
      const favorable = adjacent.some(enemy => combatOdds(state, { ...unit, x: tile.x, y: tile.y }, enemy) >= 1);
      const exposed = adjacent.length * (weak ? 7 : unit.side === 'american' ? favorable || support >= 2 ? 0.4 : 2.8 : 1.2);
      const captures = tile.value && tile.owner !== unit.side ? tile.value * 2 : 0;
      const travel = options.get(key(tile.x, tile.y))?.path.length || 0;
      return -distance(tile, target) * 2 + captures - exposed + Math.min(3, support) * 0.5 +
        (unit.side === 'american' ? terrain : terrain * 0.2) -
        (state.ruleset === 'manual' ? travel * (state.fuel < 35 && unit.side === 'german' ? 0.9 : 0.25) : 0);
    };
    const ranked = [...options.keys()].map(tileKey => state.map[tileKey]).sort((a, b) => rank(b) - rank(a));
    // Verify supply only for the best few candidates, rather than pathfinding from every reachable hex.
    const destination = ranked.slice(0, 6).find(tile => {
      if (rank(tile) <= rank(current) + 0.1) return false;
      const candidate = { ...unit, x: tile.x, y: tile.y };
      const projected = { ...state, units: state.units.map(counter => counter.id === unit.id ? candidate : counter) };
      return state.ruleset === 'manual' ? manualRules.localSupply(projected, candidate) !== 'isolated' :
        isSupplied(projected, candidate) || (!unit.supplied && !enemiesNear(state, tile, unit.side));
    });
    return destination ? { tile: destination, path: options.get(key(destination.x, destination.y)).path,
      reason: weak ? `preserve a weakened formation and seek safer ground` :
        unit.side === 'german' ? `concentrate toward ${target.name || 'the western breakthrough'} while keeping supply open` :
          `intercept ${target.name || 'the western approaches'} and cover the objectives using defensible terrain` } : null;
  }
  function planBattle(state, unit) {
    if (unit.hp < unit.maxHp * 0.3) return null;
    const targets = validTargets(state, unit).sort((a, b) => combatOdds(state, unit, b) - combatOdds(state, unit, a));
    const defender = targets.find(target => {
      const odds = combatOdds(state, unit, target);
      return odds >= (unit.side === 'german' ? 1.05 : 1.25) ||
        (target.hp <= 2 && odds >= 0.85);
    });
    return defender ? { defender, reason: unit.side === 'german' ? 'open a route with a favorable attack' : 'counterattack an exposed enemy rather than trade losses at poor odds' } : null;
  }
  function endTurn(state, humanSide = 'german') {
    if (state.ruleset === 'manual') return manualRules.endTurn(state, humanSide);
    if (state.winner || state.phase !== humanSide) return false;
    computerTurn(state, opponent(humanSide));
    return finishClassicDay(state, humanSide);
  }
  function finishClassicDay(state, humanSide) {
    updateScore(state);
    summarizeDay(state);
    const germanAlive = state.units.some(unit => unit.side === 'german' && unit.hp > 0);
    const americanAlive = state.units.some(unit => unit.side === 'american' && unit.hp > 0);
    if (!germanAlive || !americanAlive || state.turn >= LAST_TURN) {
      state.winner = germanAlive && (!americanAlive || state.score >= 8) ? 'german' : 'american';
      state.phase = 'finished';
      addLog(state, `${state.winner === 'german' ? 'GERMAN' : 'AMERICAN'} VICTORY / ${state.score} objective points.`);
      return true;
    }
    state.turn += 1;
    state.weather = state.turn < 4 ? 'Fog' : state.turn < 8 ? (random(state) < 0.6 ? 'Snow' : 'Fog') : 'Clear';
    reinforcements(state);
    beginSide(state, humanSide);
    addLog(state, `${15 + state.turn} DEC / ${state.weather.toUpperCase()} / ${humanSide === 'german' ? 'German' : 'American'} orders.`);
    state.dayStart = daySnapshot(state);
    return true;
  }
  function simulateDay(state) {
    if (state.winner) return false;
    if (state.playMode !== 'computer') throw new Error('Start a Computer vs Computer campaign to simulate a day.');
    const turn = state.turn;
    for (let step = 0; !state.winner && state.turn === turn; step++) {
      if (step > 200) throw new Error('Computer day exceeded its order limit.');
      simulateAction(state);
    }
    return true;
  }
  function simulateAction(state) {
    if (state.winner) return false;
    if (state.playMode !== 'computer') throw new Error('Start a Computer vs Computer campaign to simulate an action.');
    const manual = state.ruleset === 'manual';
    state.computerProgress ||= { stage: manual ? state.orderPhase : 'movement', index: 0 };
    const progress = state.computerProgress, side = state.phase;
    const units = state.units.filter(unit => unit.side === side);
    while (progress.index < units.length) {
      const unit = units[progress.index];
      if (progress.stage === 'movement') {
        if (manual) progress.index++;
        else progress.stage = 'combat';
        const order = planMovement(state, unit);
        if (order) {
          const from = { x: unit.x, y: unit.y };
          if (!move(state, unit.id, order.tile.x, order.tile.y)) throw new Error('Computer movement became invalid.');
          return { side, id: unit.id, path: [from, ...order.path],
            text: `${unit.name} moves to column ${order.tile.col + 1}, row ${order.tile.row + 1} to ${order.reason}.` };
        }
      } else {
        progress.index++;
        if (!manual) progress.stage = 'movement';
        if (unit.hp <= 0 || unit.exited || unit.fired) continue;
        const battle = manual ? manualRules.planBattle(state, unit) : planBattle(state, unit);
        if (battle) {
          const before = state.units.map(counter => ({ id: counter.id, hp: counter.hp, x: counter.x, y: counter.y }));
          const from = { x: unit.x, y: unit.y }, to = { x: battle.defender.x, y: battle.defender.y };
          if (manual) manualRules.resolveAttack(state, battle.ids, battle.defender.id, battle.options);
          else if (!attack(state, unit.id, battle.defender.id)) throw new Error('Computer attack became invalid.');
          const changes = before.filter(snapshot => {
            const now = state.units.find(counter => counter.id === snapshot.id);
            return now.hp !== snapshot.hp || now.x !== snapshot.x || now.y !== snapshot.y;
          }).map(snapshot => {
            const now = state.units.find(counter => counter.id === snapshot.id);
            return `${now.name}: ${snapshot.hp - now.hp} lost, ${now.hp} left${now.x !== snapshot.x || now.y !== snapshot.y ? '; repositioned (see arrows)' : ''}`;
          });
          const paths = before.filter(snapshot => {
            const now = state.units.find(counter => counter.id === snapshot.id);
            return now.x !== snapshot.x || now.y !== snapshot.y;
          }).map(snapshot => {
            const now = state.units.find(counter => counter.id === snapshot.id);
            return [{ x: snapshot.x, y: snapshot.y }, { x: now.x, y: now.y }];
          });
          return { side, id: unit.id, path: [from, to], paths, battle: true,
            text: `${unit.name}${manual && battle.ids.length > 1 ? ' and supporting formations' : ''} attacks ${battle.defender.name} to ${battle.reason}. ${changes.join('. ')}.` };
        }
      }
    }
    if (manual && progress.stage === 'movement') {
      state.orderPhase = 'combat'; progress.stage = 'combat'; progress.index = 0;
      return { side, text: `${side === 'german' ? 'Germany' : 'America'} has completed movement. Combat orders begin.` };
    }
    delete state.computerProgress;
    if (side === 'german') {
      beginSide(state, 'american');
      state.computerProgress = { stage: manual ? state.orderPhase : 'movement', index: 0 };
      return { side: 'american', text: 'German orders are complete. America now defends, repositions and counterattacks.' };
    }
    if (manual) {
      manualRules.finishDay(state);
      if (!state.winner) beginSide(state, 'german');
    } else finishClassicDay(state, 'german');
    return { side, text: state.winner ? `${state.winner.toUpperCase()} wins. Review the final situation below.` :
      `Both armies completed the day. ${15 + state.turn} December orders are ready.`, dayEnd: true };
  }
  function deserialize(text) {
    if (typeof text !== 'string' || text.length > 200000) throw new Error('Save file is too large.');
    const parsed = JSON.parse(text);
    if (parsed && parsed.version === 2) return manualRules.deserialize(parsed);
    const template = createGame();
    if (!parsed || parsed.version !== 1 || !Number.isInteger(parsed.turn) || parsed.turn < 1 || parsed.turn > LAST_TURN ||
      !['german', 'american', 'finished'].includes(parsed.phase) || !['Fog', 'Snow', 'Clear'].includes(parsed.weather) ||
      !Number.isInteger(parsed.seed) || parsed.seed < 0 || parsed.seed > 4294967295 ||
      ![null, 'german', 'american'].includes(parsed.winner) ||
      (parsed.playMode !== undefined && !['human', 'computer'].includes(parsed.playMode)) ||
      (parsed.playMode === 'computer' && !parsed.winner && parsed.phase !== 'german' && !parsed.computerProgress) ||
      (parsed.phase === 'finished') !== Boolean(parsed.winner) || !Array.isArray(parsed.units) || parsed.units.length > 30 ||
      !Array.isArray(parsed.log) || parsed.log.length > 60 || parsed.log.some(entry => typeof entry !== 'string' || entry.length > 250)) {
      throw new Error('Not a valid Winter Offensive save.');
    }
    const occupied = new Set(), ids = new Set();
    const allowedIds = new Set(['g1','g2','g3','g4','g5','g6','g7','a1','a2','a3','a4','a5','a6','a7','a8','a9','a10','a11']);
    const units = parsed.units.map(unit => {
      if (!unit || !allowedIds.has(unit.id) || ids.has(unit.id) || typeof unit.name !== 'string' || unit.name.length > 50 ||
        !['german', 'american'].includes(unit.side) || unit.side !== (unit.id.startsWith('g') ? 'german' : 'american') ||
        !['armor', 'infantry'].includes(unit.type) || !Number.isInteger(unit.x) || !Number.isInteger(unit.y) ||
        !template.map[key(unit.x, unit.y)] || typeof unit.fired !== 'boolean' || typeof unit.supplied !== 'boolean' ||
        !['hp','maxHp','attack','defense','mobility','mp'].every(field => Number.isInteger(unit[field]) && unit[field] >= 0 && unit[field] <= 12) ||
        unit.maxHp < 1 || unit.hp > unit.maxHp || unit.mp > unit.mobility || unit.attack < 1 || unit.defense < 1 || unit.mobility < 1 ||
        (unit.hp > 0 && occupied.has(key(unit.x, unit.y)))) throw new Error('Invalid unit data in save.');
      ids.add(unit.id);
      if (unit.hp > 0) occupied.add(key(unit.x, unit.y));
      return { id: unit.id, name: unit.name, side: unit.side, type: unit.type, x: unit.x, y: unit.y,
        hp: unit.hp, maxHp: unit.maxHp, attack: unit.attack, defense: unit.defense,
        mobility: unit.mobility, mp: unit.mp, fired: unit.fired, supplied: unit.supplied };
    });
    if (units.length < 14 || [...allowedIds].filter(id => !['a8','a9','a10','a11'].includes(id)).some(id => !ids.has(id))) throw new Error('Save is missing starting units.');
    for (const tile of Object.values(template.map).filter(candidate => candidate.value)) {
      const saved = parsed.map && parsed.map[key(tile.x, tile.y)];
      if (!saved || !['german', 'american'].includes(saved.owner)) throw new Error('Invalid objective data.');
      tile.owner = saved.owner;
    }
    Object.assign(template, { turn: parsed.turn, phase: parsed.phase, weather: parsed.weather,
      seed: parsed.seed, winner: parsed.winner, units, log: parsed.log });
    if (parsed.playMode !== undefined) template.playMode = parsed.playMode;
    updateScore(template);
    return restoreMetadata(parsed, template);
  }
  function describeDispatch(entry) {
    const battle = entry.match(/^(.+) > (.+): ([\d.]+):1 \/ losses (\d+):(\d+)\.$/);
    if (battle) return { kind: 'Battle', title: `${battle[1]} attacked ${battle[2]}`,
      detail: `Attacker lost ${battle[4]} strength; defender lost ${battle[5]} strength. Attack odds were ${battle[3]} to 1.` };
    const manualBattle = entry.match(/^(.+) > (.+): ([\d.]+):1 \/ (\w+) vs (\w+)\.$/);
    if (manualBattle) return { kind: 'Battle', title: `${manualBattle[1].split(' + ').join(' and ')} attacked ${manualBattle[2]}`,
      detail: `Attack odds were ${manualBattle[3]} to 1. Attack strategy: ${manualRules.strategyLabel(manualBattle[4])}. Defense strategy: ${manualRules.strategyLabel(manualBattle[5])}.` };
    const day = entry.match(/^(\d+) DEC \/ ([A-Z]+) \/ (.+)\.$/);
    if (day) return { kind: 'New day', title: `${day[1]} December — ${day[2].toLowerCase()} weather`,
      detail: `${day[3]}.` };
    const victory = entry.match(/^(GERMAN|AMERICAN) VICTORY \/ (.+)\.$/);
    if (victory) return { kind: 'Result', title: `${victory[1] === 'GERMAN' ? 'Germany' : 'America'} won the campaign`,
      detail: `Final German score: ${victory[2]}.` };
    const capture = entry.match(/^(.+) captures (.+)\.$/);
    if (capture) return { kind: 'Objective', title: `${capture[1]} captured ${capture[2]}`,
      detail: 'Control of this objective changed sides.' };
    const arrival = entry.match(/^(.+) arrives\.$/);
    if (arrival) return { kind: 'Reinforcement', title: `${arrival[1]} arrived`,
      detail: 'This formation has joined the battlefield.' };
    const eliminated = entry.match(/^(.+) eliminated\.$/);
    if (eliminated) return { kind: 'Elimination', title: `${eliminated[1]} was eliminated`,
      detail: 'The formation has no strength remaining and is removed from play.' };
    const retreat = entry.match(/^(.+) retreats\.$/);
    if (retreat) return { kind: 'Retreat', title: `${retreat[1]} retreated`,
      detail: 'The defender withdrew to another hex after combat.' };
    const blocked = entry.match(/^(.+): retreat blocked\.$/);
    if (blocked) return { kind: 'Retreat', title: `${blocked[1]} could not retreat`,
      detail: 'Its retreat was blocked, causing additional strength losses.' };
    const supply = entry.match(/^(.+): general supply failed; strength halved\.$/);
    if (supply) return { kind: 'Supply', title: `${supply[1]} did not receive combat supplies`,
      detail: 'Its effective strength was halved for this battle, not permanently.' };
    return { kind: /supply|airdrop/i.test(entry) ? 'Supply' : /retreat|advance/i.test(entry) ? 'Maneuver' :
      /entry blocked|arrival/i.test(entry) ? 'Reinforcement' : 'Update', title: entry, detail: '' };
  }
  function campaignStanding(state) {
    const target = 8;
    const leader = state.winner || (state.score >= target ? 'german' : 'american');
    return { points: state.score, target, leader,
      status: state.winner ? `${leader.toUpperCase()} VICTORY` : `${leader.toUpperCase()} ON COURSE`,
      detail: state.winner ? 'Campaign finished.' :
        state.score >= target ? 'German target reached; hold it through the final victory check.' :
          `Germany needs ${target - state.score} more points to reach its target.`,
      provisional: !state.winner };
  }
  const api = { createGame, neighbors, occupant, distance, reachable, move, attack, combatOdds,
    isSupplied, beginSide, endTurn, computerTurn, simulateDay, simulateAction, deserialize, arrivalsThisTurn, campaignStanding, describeDispatch, armySummary, reportOperation, TERRAIN, key,
    battlePreview, orderReason, validTargets, parseSeed, restoreMetadata, daySnapshot, summarizeDay, planMovement };
  manualRules = createManualRules(api);
  api.manual = manualRules;
  if (typeof module === 'object') module.exports = api;
  root.WinterGame = api;
  if (typeof document === 'undefined') return;

  const canvas = document.getElementById('battlefield');
  const context = canvas.getContext('2d');
  const $ = id => document.getElementById(id);
  const freshSeed = () => crypto.getRandomValues(new Uint32Array(1))[0];
  let state = createGame(freshSeed()), commandSide = 'german', selected = 'g1', mode = 'move', busy = false;
  let actionVisual = null;
  let visualStarted = 0, animationFrame = null;
  function setActionVisual(value) {
    actionVisual = value; visualStarted = performance.now();
    if (animationFrame !== null) cancelAnimationFrame(animationFrame);
    animationFrame = null;
    if (!value?.path || window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    const frame = () => {
      draw();
      if (performance.now() - visualStarted < 1200) animationFrame = requestAnimationFrame(frame);
      else animationFrame = null;
    };
    animationFrame = requestAnimationFrame(frame);
  }
  let cursor = { x: state.units[0].x, y: state.units[0].y };
  const SAVE_KEY = 'winter-offensive-v1';
  const COMMAND_SIDE_KEY = 'winter-offensive-command-side';
  const moveHistory = [];
  function resetSideReports() {
    setActionVisual(null);
    $('action-story').textContent = 'Campaign ready. Select a formation or follow computer orders with Next Action.';
  }
  function renderSideReports() {
    for (const side of ['american', 'german']) {
      const summary = armySummary(state, side);
      $(`${side}-army-summary`).textContent = `${summary.formations} active formations / ${summary.strength} ${state.ruleset === 'manual' ? 'combat points' : 'strength'} / ${summary.isolated} isolated`;
      $(`${side}-objectives`).textContent = summary.objectives.length ? `Holding: ${summary.objectives.join(', ')}.` : 'Holding no objectives.';
      const list = $(`${side}-reports`);
      list.replaceChildren();
      const events = state.reports?.[side] || [];
      if (!events.length) {
        const item = document.createElement('li');
        item.textContent = 'No operations recorded yet. The totals above show the current position.';
        list.append(item);
      }
      for (const event of events) {
        const item = document.createElement('li');
        item.textContent = `${15 + event.turn} DEC: ${event.text}`;
        list.append(item);
      }
    }
  }
  let autoplay = false, autoplayTimer = null;
  const observing = () => state.playMode === 'computer';
  function pauseSimulation() {
    autoplay = false;
    if (autoplayTimer !== null) clearTimeout(autoplayTimer);
    autoplayTimer = null;
  }
  function message(text) { $('notice').textContent = text; }
  function renderBattleUnit(id, snapshot, unit) {
    const panel = $(id);
    panel.dataset.side = snapshot.side;
    panel.querySelector('h3').textContent = snapshot.name;
    const strength = panel.querySelector('.battle-strength');
    const segments = state.ruleset === 'manual' ? Math.ceil(unit.hp / Math.max(1, snapshot.hp) * 12) : unit.hp;
    strength.replaceChildren(...Array.from({ length: segments }, () => document.createElement('i')));
    strength.setAttribute('aria-label', `${snapshot.name}: ${unit.hp} strength remaining out of ${snapshot.hp}.`);
    const losses = snapshot.hp - unit.hp;
    panel.querySelector('.battle-loss').textContent = `${snapshot.hp} → ${unit.hp} / ${losses ? `-${losses}` : 'no losses'}`;
  }
  function showBattleReport(attacker, defender, odds) {
    $('battle-details').replaceChildren();
    const attackerNow = state.units.find(unit => unit.id === attacker.id);
    const defenderNow = state.units.find(unit => unit.id === defender.id);
    renderBattleUnit('battle-attacker', attacker, attackerNow);
    renderBattleUnit('battle-defender', defender, defenderNow);
    $('battle-odds').textContent = `${odds.toFixed(1)}:1 attack odds`;
    const defenderMoved = defenderNow.hp > 0 && (defenderNow.x !== defender.x || defenderNow.y !== defender.y);
    $('battle-outcome').textContent = defenderNow.hp === 0 ? `${defender.name} eliminated.` :
      attackerNow.hp === 0 ? `${attacker.name} eliminated.` :
      defenderMoved ? `${defender.name} retreats.` : 'Both formations hold their positions.';
    setActionVisual({ side: attacker.side, path: [attacker, defender], battle: true,
      paths: defenderMoved ? [[defender, defenderNow]] : [] });
    $('action-story').textContent = `${attacker.name} attacked ${defender.name}: attacker lost ${attacker.hp - attackerNow.hp}, defender lost ${defender.hp - defenderNow.hp}. ${$('battle-outcome').textContent}`;
    draw();
    $('battle-dialog').showModal();
  }
  function persist() {
    try { localStorage.setItem(SAVE_KEY, JSON.stringify(state)); }
    catch { message('Browser storage unavailable. Export a save to keep your campaign.'); }
  }
  try {
    const savedSide = localStorage.getItem(COMMAND_SIDE_KEY);
    if (savedSide === 'german' || savedSide === 'american') commandSide = savedSide;
    const saved = localStorage.getItem(SAVE_KEY);
    if (saved) {
      state = deserialize(saved);
      if (state.ruleset === 'manual') commandSide = state.playerSide;
      if (observing()) commandSide = 'german';
      selected = state.units.find(unit => unit.side === commandSide && unit.hp > 0 && !unit.exited)?.id || null;
      if (state.ruleset === 'manual') mode = state.orderPhase === 'combat' ? 'attack' : 'move';
    }
  } catch { message('Saved campaign could not be loaded. A new campaign is ready.'); }
  const restoredSelection = state.units.find(unit => unit.id === selected);
  if (restoredSelection) cursor = { x: restoredSelection.x, y: restoredSelection.y };
  const center = tile => ({ x: 30 + tile.x * 23, y: 31 + tile.y * 40 });
  function hex(tile, inset = 0) {
    const point = center(tile);
    context.beginPath();
    for (let index = 0; index < 6; index++) {
      const angle = (index * 60 - 30) * Math.PI / 180;
      const x = point.x + (26 - inset) * Math.cos(angle), y = point.y + (26 - inset) * Math.sin(angle);
      if (index === 0) context.moveTo(x, y); else context.lineTo(x, y);
    }
    context.closePath();
  }
  function draw() {
    const style = getComputedStyle(document.documentElement);
    const dark = document.documentElement.dataset.theme === 'dark';
    const ink = style.getPropertyValue(dark ? '--cp-text' : '--cp-surface').trim();
    const paper = style.getPropertyValue(dark ? '--cp-surface' : '--cp-text').trim();
    const palette = {
      plain: ink,
      forest: style.getPropertyValue('--cp-success').trim(),
      ridge: style.getPropertyValue('--cp-warning').trim(),
      river: style.getPropertyValue('--cp-link').trim(),
      road: style.getPropertyValue('--cp-warning').trim(),
      town: style.getPropertyValue('--cp-accent').trim(),
      german: style.getPropertyValue('--cp-danger').trim(),
      american: style.getPropertyValue('--cp-link').trim(),
      highlight: style.getPropertyValue('--cp-warning').trim(),
      counterText: style.getPropertyValue(dark ? '--cp-accent-fg' : '--cp-surface').trim()
    };
    context.fillStyle = paper; context.fillRect(0, 0, canvas.width, canvas.height);
    context.font = '12px Consolas, monospace'; context.textAlign = 'center'; context.textBaseline = 'middle';
    const unit = state.units.find(candidate => candidate.id === selected);
    const options = mode === 'move' && !observing() ? reachable(state, unit) : new Map();
    for (const tile of Object.values(state.map)) {
      const point = center(tile);
      context.fillStyle = palette[tile.terrain];
      context.globalAlpha = tile.value ? 0.5 : tile.terrain === 'plain' ? 0.04 : tile.terrain === 'river' ? 0.4 : 0.23;
      hex(tile, 1); context.fill(); context.globalAlpha = 1;
      context.strokeStyle = ink; context.globalAlpha = 0.25; context.lineWidth = 1;
      hex(tile); context.stroke(); context.globalAlpha = 1;
      if (options.has(key(tile.x, tile.y))) {
        context.fillStyle = paper; context.beginPath();
        context.arc(point.x, point.y + 8, 5, 0, Math.PI * 2); context.fill();
        context.fillStyle = ink; context.beginPath();
        context.arc(point.x, point.y + 8, 3, 0, Math.PI * 2); context.fill();
      }
      context.fillStyle = tile.value ? ink : palette[tile.terrain];
      context.globalAlpha = tile.value ? 1 : 0.9;
      context.fillText(TERRAIN[tile.terrain].symbol, point.x, point.y);
      context.globalAlpha = 1;
    }
    for (const counter of state.units.filter(candidate => candidate.hp > 0 && !candidate.exited)) {
      const point = center(counter);
      context.fillStyle = palette[counter.side];
      context.strokeStyle = ink; context.lineWidth = 1;
      context.fillRect(point.x - 16, point.y - 12, 32, 24); context.strokeRect(point.x - 16, point.y - 12, 32, 24);
      context.fillStyle = palette.counterText;
      context.font = 'bold 11px Consolas, monospace';
      context.fillText(`${counter.side === 'german' ? 'G' : 'A'}${counter.type === 'armor' ? '=' : 'X'}`, point.x, point.y - 4);
      context.font = '10px Consolas, monospace'; context.fillText(`${counter.hp}/${counter.mp}`, point.x, point.y + 7);
      if (counter.id === selected) { context.strokeStyle = palette.highlight; context.lineWidth = 2; hex(counter, 1); context.stroke(); }
      if (!counter.supplied) { context.fillStyle = palette.highlight; context.fillText('!', point.x + 20, point.y); }
      if (!observing() && mode === 'attack' && validTargets(state, unit).includes(counter)) {
        context.strokeStyle = palette.german; context.lineWidth = 2; context.strokeRect(point.x - 20, point.y - 16, 40, 32);
      }
    }
    for (const tile of Object.values(state.map).filter(candidate => candidate.value)) {
      const point = center(tile);
      context.strokeStyle = palette.town; context.lineWidth = 2;
      hex(tile); context.stroke(); hex(tile, 3); context.stroke();
      context.font = 'bold 10px Consolas, monospace';
      const labelWidth = context.measureText(tile.name).width + 8;
      context.fillStyle = paper;
      context.fillRect(point.x - labelWidth / 2, point.y + 13, labelWidth, 13);
      context.fillRect(point.x - 14, point.y - 25, 28, 13);
      context.fillStyle = ink;
      context.fillText(tile.name, point.x, point.y + 20);
      context.fillStyle = palette[tile.owner];
      context.fillText(`${tile.owner === 'german' ? 'G' : 'A'}${tile.value}`, point.x, point.y - 18);
    }
    // Draw arrivals last so objective labels cannot hide their markers.
    for (const counter of arrivalsThisTurn(state)) {
      const point = center(counter);
      context.strokeStyle = palette.highlight; context.lineWidth = 3;
      context.strokeRect(point.x - 18, point.y - 14, 36, 28);
      context.fillStyle = paper;
      context.fillRect(point.x - 15, point.y - 27, 30, 12);
      context.fillStyle = palette.highlight; context.font = 'bold 10px Consolas, monospace';
      context.fillText('NEW', point.x, point.y - 21);
    }
    if (actionVisual?.path) {
      const arrow = (path, color, dashed = false) => {
        if (path.length < 2) return;
        const points = path.map(center);
        context.strokeStyle = color; context.fillStyle = color; context.lineWidth = 4;
        context.setLineDash(dashed ? [7, 5] : []);
        context.beginPath(); context.moveTo(points[0].x, points[0].y);
        for (const point of points.slice(1)) context.lineTo(point.x, point.y);
        context.stroke(); context.setLineDash([]);
        const end = points[points.length - 1], previous = points[points.length - 2];
        const angle = Math.atan2(end.y - previous.y, end.x - previous.x);
        context.beginPath(); context.moveTo(end.x, end.y);
        context.lineTo(end.x - 12 * Math.cos(angle - 0.5), end.y - 12 * Math.sin(angle - 0.5));
        context.lineTo(end.x - 12 * Math.cos(angle + 0.5), end.y - 12 * Math.sin(angle + 0.5));
        context.closePath(); context.fill();
      };
      arrow(actionVisual.path, palette[actionVisual.side], actionVisual.battle);
      for (const path of actionVisual.paths || []) arrow(path, palette.highlight);
      if (!actionVisual.preview && performance.now() - visualStarted < 1200) {
        const paths = actionVisual.battle ? actionVisual.paths || [] : [actionVisual.path];
        const fraction = Math.min(1, (performance.now() - visualStarted) / 1000);
        for (const path of paths.filter(path => path.length >= 2)) {
          const progress = fraction * (path.length - 1), index = Math.min(path.length - 2, Math.floor(progress));
          const start = center(path[index]), end = center(path[index + 1]), amount = progress - index;
          context.fillStyle = palette.highlight; context.strokeStyle = ink; context.lineWidth = 2;
          context.beginPath(); context.arc(start.x + (end.x - start.x) * amount, start.y + (end.y - start.y) * amount, 9, 0, Math.PI * 2);
          context.fill(); context.stroke();
        }
      }
    }
    const cursorTile = state.map[key(cursor.x, cursor.y)];
    if (cursorTile) { context.strokeStyle = ink; context.globalAlpha = 0.7; context.lineWidth = 1; hex(cursorTile, 5); context.stroke(); context.globalAlpha = 1; }
    context.fillStyle = ink; context.textAlign = 'left'; context.font = '11px Consolas, monospace';
    context.fillText('WEST / MEUSE', 10, 499); context.textAlign = 'right'; context.fillText('EAST / GERMAN SUPPLY', 720, 499);
  }
  function render() {
    draw();
    renderSideReports();
    $('date').textContent = `${15 + state.turn} DEC 1944`;
    $('turn').textContent = `TURN ${state.turn} / ${LAST_TURN}`;
    $('weather').textContent = state.weather.toUpperCase();
    $('seed-status').textContent = state.initialSeed === undefined ? 'LEGACY / SEED UNKNOWN' : `SEED ${state.initialSeed}`;
    $('score').textContent = `${state.score} / 8`;
    const manual = state.ruleset === 'manual';
    for (const side of ['american', 'german']) {
      const summary = armySummary(state, side);
      $(`${side}-held-points`).textContent = String(summary.objectivePoints);
      $(`${side}-score-formations`).textContent = String(summary.formations);
      $(`${side}-score-strength`).textContent = String(summary.strength);
      $(`${side}-score-isolated`).textContent = String(summary.isolated);
      $(`${side}-strength-label`).textContent = manual ? 'Combat points remaining' : 'Strength remaining';
    }
    const standing = campaignStanding(state);
    $('campaign-points').textContent = `${standing.points} / ${standing.target}`;
    $('campaign-leader').textContent = standing.status;
    $('campaign-leader').dataset.side = standing.leader;
    $('campaign-score-detail').textContent = `${standing.detail}${standing.provisional ? ' Provisional: combat, captures and elimination can change the outcome.' : ''}`;
    $('campaign-score-basis').textContent = manual ?
      'Adapted German score: objectives + exits − Allied breakthrough. Not the original SSI scoring table.' :
      'German score: currently controlled objective points.';
    $('campaign-score-progress').value = Math.min(standing.target, standing.points);
    $('phase').textContent = state.winner ? `${state.winner.toUpperCase()} VICTORY` : busy ? 'PROCESSING PHASE' :
      observing() ? `COMPUTER VS COMPUTER / ${autoplay ? 'PLAYING' : 'PAUSED'}` :
        `${commandSide.toUpperCase()} ${manual ? state.orderPhase.toUpperCase() : 'ORDERS'}`;
    $('spectator-controls').hidden = !observing();
    $('simulation-play').textContent = autoplay ? 'Pause' : 'Play';
    $('simulation-play').disabled = Boolean(state.winner) || busy;
    $('simulation-step').disabled = Boolean(state.winner) || busy || autoplay;
    $('simulation-action').disabled = Boolean(state.winner) || busy || autoplay;
    for (const side of ['american', 'german']) {
      $(`${side}-day-summary`).textContent = state.daySummary?.[side] ||
        `${side === 'american' ? 'America' : 'Germany'}: the first day has not finished yet.`;
    }
    $('ruleset-status').textContent = manual ? 'MANUAL-BASED ADAPTATION' : 'CLASSIC RULES';
    $('manual-command').hidden = !manual;
    if (manual) {
      $('resources').textContent = `Supply ${state.generalSupply[commandSide]} / Artillery ${state.artillery[commandSide]} / German fuel ${state.fuel}`;
      $('defense-strategy').value = state.defenseStrategy;
      $('defense-strategy').disabled = observing() || busy;
    }
    $('end-turn').textContent = manual && state.orderPhase === 'movement' ? 'End Movement' : manual ? 'End Combat / Day' : 'End Turn';
    const unit = state.units.find(candidate => candidate.id === selected);
    $('unit-name').textContent = unit ? unit.name : 'No unit selected';
    $('unit-kind').textContent = unit ? `${unit.side.toUpperCase()} / ${unit.type.toUpperCase()}` : 'GERMAN COMMAND';
    $('strength').textContent = unit ? `${unit.hp} / ${unit.maxHp}` : '-';
    $('movement').textContent = unit ? String(unit.mp) : '-';
    $('supply').textContent = unit ? manual ? manualRules.localSupply(state, unit).toUpperCase() : unit.supplied ? 'CONNECTED' : 'ISOLATED' : '-';
    $('supply').dataset.status = unit ? unit.supplied ? 'connected' : 'isolated' : '';
    $('combat').textContent = unit ? `${unit.fired ? 'COMMITTED' : 'AVAILABLE'}${manual ? ` / readiness ${unit.readiness}/7` : ''}` : '-';
    for (const command of ['move', 'attack', 'end-turn']) {
      $(command).disabled = observing() || busy || Boolean(state.winner) || (command !== 'end-turn' && (!unit || unit.hp <= 0 || unit.exited)) ||
        (manual && command !== 'end-turn' && (command === 'move' ? state.orderPhase !== 'movement' : state.orderPhase !== 'combat'));
      if (command !== 'end-turn') {
        const reason = orderReason(state, unit, command);
        $(command).disabled ||= Boolean(reason);
        $(command).title = reason || (command === 'move' ? 'Show reachable destinations.' : 'Show eligible targets; click one to preview combat.');
      }
    }
    const moveReason = orderReason(state, unit, 'move'), attackReason = orderReason(state, unit, 'attack');
    const targets = validTargets(state, unit);
    $('order-guidance').textContent = observing() ? moveReason :
      `Move: ${moveReason || 'click a dotted destination; movement is allowed.'} Attack: ${attackReason ||
        (targets.length ? `eligible targets: ${targets.map(target => target.name).join(', ')}. Choose Attack, then click a target for a risk preview.` :
          'no adjacent eligible enemy. Move closer or end the current phase.')}`;
    $('move').setAttribute('aria-pressed', String(mode === 'move'));
    $('attack').setAttribute('aria-pressed', String(mode === 'attack'));
    $('undo-move').disabled = observing() || busy || Boolean(state.winner) || moveHistory.length === 0;
    $('roster').replaceChildren();
    state.units.filter(candidate => (observing() || candidate.side === commandSide) && candidate.hp > 0 && !candidate.exited).forEach(counter => {
      const button = document.createElement('button');
      button.className = 'roster-unit'; button.setAttribute('aria-pressed', String(counter.id === selected));
      button.textContent = `${observing() ? `${counter.side === 'german' ? 'G' : 'A'} / ` : ''}${counter.type === 'armor' ? '[=]' : '[X]'} ${counter.name}`;
      const stats = document.createElement('span'); stats.textContent = `${counter.hp} / ${counter.mp}`; button.append(stats);
      button.disabled = busy;
      button.onclick = () => { selected = counter.id; cursor = { x: counter.x, y: counter.y }; render(); };
      $('roster').append(button);
    });
    $('objectives').replaceChildren();
    Object.values(state.map).filter(tile => tile.value).forEach(tile => {
      const row = document.createElement('li');
      row.dataset.side = tile.owner;
      row.textContent = `${tile.name} / ${tile.owner === 'german' ? 'GER' : 'USA'} / ${tile.value} PT — ${tile.name === 'MEUSE' || tile.name === 'DINANT' ? 'western crossing / breakthrough objective' : 'crossroads / defensive anchor'}. Capture to transfer ${tile.value} held points.`;
      $('objectives').append(row);
    });
    $('dispatches').replaceChildren();
    $('history-count').textContent = `${state.log.length} saved events`;
    state.log.forEach(entry => {
      const event = describeDispatch(entry);
      const line = document.createElement('li');
      const label = document.createElement('span'); label.className = 'history-label'; label.textContent = event.kind;
      const content = document.createElement('div');
      const title = document.createElement('strong'); title.textContent = event.title; content.append(title);
      if (event.detail) {
        const detail = document.createElement('p'); detail.textContent = event.detail; content.append(detail);
      }
      line.append(label, content); $('dispatches').append(line);
    });
    const arrivals = arrivalsThisTurn(state);
    $('reinforcement-arrivals').hidden = arrivals.length === 0;
    $('arrival-list').replaceChildren();
    for (const counter of arrivals) {
      const tile = state.map[key(counter.x, counter.y)];
      const button = document.createElement('button');
      button.textContent = `${counter.side.toUpperCase()} / ${counter.name} / column ${tile.col + 1}, row ${tile.row + 1}`;
      button.disabled = busy;
      button.onclick = () => {
        cursor = { x: counter.x, y: counter.y };
        if (observing() || counter.side === commandSide) selected = counter.id;
        render();
        canvas.focus();
        canvas.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      };
      $('arrival-list').append(button);
    }
    $('new-game').disabled = busy; $('import').disabled = busy; $('export').disabled = busy;
  }
  function activate(tile) {
    if (!tile || busy || state.winner) return;
    const target = occupant(state, tile);
    if (observing()) {
      if (target) selected = target.id;
      render(); return;
    }
    const unit = state.units.find(candidate => candidate.id === selected);
    if (target && target.side === commandSide) { selected = target.id; render(); return; }
    if (mode === 'attack' && target && unit) {
      openCombatPlan(target); return;
    } else if (mode === 'move') {
      const previous = { save: JSON.stringify(state), unitId: selected };
      const option = reachable(state, unit).get(key(tile.x, tile.y));
      const from = unit ? { x: unit.x, y: unit.y } : null;
      if (move(state, selected, tile.x, tile.y)) {
        setActionVisual({ side: unit.side, path: [from, ...option.path] });
        $('action-story').textContent = `${unit.name} moved along the highlighted route to column ${tile.col + 1}, row ${tile.row + 1}. ${unit.mp} movement points remain.`;
        moveHistory.push(previous); message('Movement complete.'); persist();
      } else message(orderReason(state, unit, 'move') || 'Destination blocked: choose an empty dotted hex within the movement budget. Enemy zones and terrain restrict the route.');
    } else message('No enemy target in this hex.');
    render();
  }
  canvas.addEventListener('click', event => {
    const bounds = canvas.getBoundingClientRect();
    const x = (event.clientX - bounds.left) * canvas.width / bounds.width;
    const y = (event.clientY - bounds.top) * canvas.height / bounds.height;
    const tile = Object.values(state.map).sort((first, second) => {
      const firstPoint = center(first), secondPoint = center(second);
      return Math.hypot(firstPoint.x - x, firstPoint.y - y) - Math.hypot(secondPoint.x - x, secondPoint.y - y);
    })[0];
    const point = center(tile);
    if (Math.hypot(point.x - x, point.y - y) < 27) { cursor = { x: tile.x, y: tile.y }; activate(tile); }
  });
  canvas.addEventListener('keydown', event => {
    const delta = { ArrowLeft: [-2, 0], ArrowRight: [2, 0], ArrowUp: [cursor.y % 2 ? -1 : 1, -1], ArrowDown: [cursor.y % 2 ? -1 : 1, 1] }[event.key];
    if (delta) {
      event.preventDefault();
      const tile = state.map[key(cursor.x + delta[0], cursor.y + delta[1])];
      if (tile) { cursor = { x: tile.x, y: tile.y }; render(); }
    } else if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); activate(state.map[key(cursor.x, cursor.y)]); }
  });
  $('move').onclick = () => { mode = 'move'; render(); };
  $('attack').onclick = () => { mode = 'attack'; render(); };
  let plannedDefender = null;
  function openCombatPlan(defender) {
    const unit = state.units.find(candidate => candidate.id === selected);
    const manual = state.ruleset === 'manual';
    const eligible = manual ? manualRules.eligibleAttackers(state, defender) : validTargets(state, unit).includes(defender) ? [unit] : [];
    if (!eligible.length) { message(orderReason(state, unit, 'attack') || 'No eligible attackers: the enemy must be adjacent and not already defended in this manual combat phase.'); return; }
    plannedDefender = defender.id;
    $('plan-target').textContent = `Attack ${defender.name} / ${defender.hp} ${manual ? 'CP' : 'strength'}`;
    $('manual-plan-options').hidden = !manual;
    $('plan-attackers').replaceChildren();
    for (const unit of eligible) {
      const label = document.createElement('label'), input = document.createElement('input');
      input.type = 'checkbox'; input.value = unit.id; input.checked = unit.id === selected;
      label.append(input, document.createTextNode(` ${unit.name} / ${unit.hp} ${manual ? `CP / readiness ${unit.readiness}` : 'strength'}`));
      $('plan-attackers').append(label);
    }
    if (!$('plan-attackers').querySelector('input:checked')) $('plan-attackers').querySelector('input').checked = true;
    $('plan-artillery').value = '0';
    $('plan-artillery').max = manual ? String(Math.min(eligible.length, state.artillery[commandSide])) : '0';
    $('plan-error').textContent = '';
    const previewUnit = eligible.find(attacker => attacker.id === selected) || eligible[0];
    setActionVisual({ side: previewUnit.side, path: [previewUnit, defender], battle: true, preview: true });
    draw(); updatePreview();
    $('combat-plan').showModal();
  }
  function updatePreview() {
    try {
      const ids = [...$('plan-attackers').querySelectorAll('input:checked')].map(input => input.value);
      const preview = battlePreview(state, ids, plannedDefender, {
        strategy: $('plan-strategy').value, artillery: Number($('plan-artillery').value)
      });
      const heading = document.createElement('strong'); heading.textContent = `${preview.odds.toFixed(2)}:1 nominal attack odds`;
      const detail = document.createElement('p'); detail.textContent = preview.detail;
      const risk = document.createElement('p'); risk.textContent = preview.risk;
      $('plan-preview').replaceChildren(heading, detail, risk);
      $('resolve-attack').disabled = false; $('plan-error').textContent = '';
    } catch (error) {
      $('plan-preview').textContent = error.message; $('resolve-attack').disabled = true;
    }
  }
  $('combat-plan').addEventListener('input', updatePreview);
  $('combat-plan').addEventListener('change', updatePreview);
  $('combat-plan').addEventListener('close', () => { if (actionVisual?.preview) setActionVisual(null); draw(); });
  $('resolve-attack').onclick = () => {
    try {
      const ids = [...$('plan-attackers').querySelectorAll('input:checked')].map(input => input.value);
      const attacker = { ...state.units.find(unit => unit.id === ids[0]) };
      const defenderBefore = { ...state.units.find(unit => unit.id === plannedDefender) };
      const positionsBefore = state.units.map(unit => ({ ...unit }));
      if (state.ruleset !== 'manual') {
        const odds = battlePreview(state, ids, plannedDefender).odds;
        if (!attack(state, ids[0], plannedDefender)) throw new Error('This attack is no longer eligible.');
        moveHistory.length = 0; persist(); $('combat-plan').close(); render();
        showBattleReport(attacker, defenderBefore, odds); message('Battle resolved. Review the actual losses.');
        return;
      }
      const result = manualRules.resolveAttack(state, ids, plannedDefender, {
        strategy: $('plan-strategy').value, artillery: Number($('plan-artillery').value)
      });
      moveHistory.length = 0; persist(); $('combat-plan').close(); render();
      const attackers = result.units.slice(0, -1), defender = result.units[result.units.length - 1];
      renderBattleUnit('battle-attacker', { side: commandSide, name: `${attackers.length} attacking formation(s)`,
        hp: attackers.reduce((sum, unit) => sum + unit.hp, 0) }, { hp: attackers.reduce((sum, unit) => sum + unit.remaining, 0) });
      renderBattleUnit('battle-defender', defender, { hp: defender.remaining });
      $('battle-odds').textContent = `${result.odds.toFixed(1)}:1 / ${manualRules.strategyLabel(result.attackStrategy)} vs ${manualRules.strategyLabel(result.defenseStrategy)} / artillery ${result.artillery}:${result.defensiveArtillery}`;
      $('battle-details').replaceChildren();
      for (const unit of result.units) {
        const item = document.createElement('li'); item.textContent = `${unit.name}: ${unit.hp} → ${unit.remaining} CP (loss ${unit.loss})`;
        $('battle-details').append(item);
      }
      $('battle-outcome').textContent = result.events.join(' ') || 'Both formations hold their positions.';
      setActionVisual({ side: commandSide, path: [attacker, defenderBefore], battle: true,
        paths: [...result.advancePaths.map(entry => {
          return [positionsBefore.find(unit => unit.id === entry.id), ...entry.path.map(tileKey => state.map[tileKey])];
        }), ...(result.retreatPath.length ? [[defenderBefore, ...result.retreatPath.map(tileKey => state.map[tileKey])]] : []),
          ...result.attackerRetreats.map(entry => [positionsBefore.find(unit => unit.id === entry.id), ...entry.path.map(tileKey => state.map[tileKey])])] });
      $('action-story').textContent = `Battle complete: ${result.units.map(unit => `${unit.name} lost ${unit.loss} CP (${unit.remaining} left)`).join('; ')}. ${$('battle-outcome').textContent}`;
      draw();
      $('battle-dialog').showModal(); message('Battle resolved. Review the graphical report.');
    } catch (error) { $('plan-error').textContent = error.message; }
  };
  $('defense-strategy').onchange = event => {
    if (state.ruleset !== 'manual' || busy || observing()) return;
    state.defenseStrategy = event.target.value; persist();
  };
  $('undo-move').onclick = () => {
    if (observing() || busy || state.winner || moveHistory.length === 0) return;
    const previous = moveHistory[moveHistory.length - 1];
    try {
      const restored = deserialize(previous.save);
      resetSideReports();
      state = restored; moveHistory.pop(); selected = previous.unitId; mode = 'move';
      const unit = state.units.find(candidate => candidate.id === selected);
      cursor = { x: unit.x, y: unit.y };
      persist(); message(`Move undone: ${unit.name}.`); render();
    } catch (error) { message(`Move could not be undone: ${error.message}`); }
  };
  $('end-turn').onclick = () => {
    if (observing() || busy || state.winner) return;
    moveHistory.length = 0;
    busy = true; render();
    setTimeout(() => {
      try {
        endTurn(state, commandSide);
        if (state.ruleset === 'manual') mode = state.orderPhase === 'combat' ? 'attack' : 'move';
        if (!state.units.some(unit => unit.id === selected && unit.hp > 0 && !unit.exited)) {
          selected = state.units.find(unit => unit.side === commandSide && unit.hp > 0 && !unit.exited)?.id || null;
        }
        persist(); message(state.winner ? 'Campaign complete.' : 'New orders ready.');
      }
      catch (error) { message(`Turn could not complete: ${error.message}`); }
      finally { busy = false; render(); }
    }, 80);
  };
  function runSimulationDay(wholeDay = false) {
    if (!observing() || busy || state.winner) { pauseSimulation(); render(); return; }
    busy = true; render();
    try {
      if (wholeDay) {
        simulateDay(state); setActionVisual(null);
        $('action-story').textContent = 'The remaining orders today are complete. Read the end-of-day summary and side reports below.';
      } else {
        const action = simulateAction(state);
        setActionVisual(action);
        if (action.id) selected = action.id;
        $('action-story').textContent = action.text;
      }
      if (!state.units.some(unit => unit.id === selected && unit.hp > 0 && !unit.exited)) {
        selected = state.units.find(unit => unit.hp > 0 && !unit.exited)?.id || null;
      }
      persist();
      message(state.winner ? 'Computer campaign complete.' : wholeDay ? 'Day complete. Review the situation below.' : 'One computer order resolved. Follow its highlighted route or battle arrow.');
      if (state.winner) pauseSimulation();
    } catch (error) {
      pauseSimulation(); message(`Simulation paused: ${error.message}`);
    } finally { busy = false; render(); }
    if (autoplay) autoplayTimer = setTimeout(() => runSimulationDay(), Number($('simulation-speed').value));
  }
  $('simulation-play').onclick = () => {
    if (!observing() || busy || state.winner) return;
    if (autoplay) { pauseSimulation(); render(); }
    else { autoplay = true; runSimulationDay(); }
  };
  $('simulation-step').onclick = () => {
    if (!autoplay) runSimulationDay(true);
  };
  $('simulation-action').onclick = () => { if (!autoplay) runSimulationDay(); };
  $('simulation-speed').onchange = () => {
    if (autoplay) {
      clearTimeout(autoplayTimer);
      autoplayTimer = setTimeout(() => runSimulationDay(), Number($('simulation-speed').value));
    }
  };
  $('new-game').onclick = () => { pauseSimulation(); render(); $('restart-error').textContent = ''; $('restart-dialog').showModal(); };
  document.querySelectorAll('[data-command-side]').forEach(button => button.onclick = () => {
    pauseSimulation();
    let seed;
    try {
      if ($('campaign-seed').validity.badInput) throw new Error('Enter a valid whole-number seed or clear the field for a random campaign.');
      seed = $('campaign-seed').value === '' ? freshSeed() : parseSeed($('campaign-seed').value);
    }
    catch (error) { $('restart-error').textContent = error.message; return; }
    const computer = button.dataset.commandSide === 'computer';
    commandSide = computer ? 'german' : button.dataset.commandSide;
    moveHistory.length = 0;
    state = createGame(seed, $('campaign-ruleset').value, commandSide);
    resetSideReports();
    if (computer) state.playMode = 'computer';
    if (commandSide === 'american' && state.ruleset !== 'manual') {
      computerTurn(state, 'german');
      beginSide(state, 'american');
    }
    selected = state.units.find(unit => unit.side === commandSide && unit.hp > 0)?.id || null;
    const unit = state.units.find(candidate => candidate.id === selected);
    mode = state.ruleset === 'manual' && state.orderPhase === 'combat' ? 'attack' : 'move';
    cursor = { x: unit.x, y: unit.y };
    $('restart-dialog').close(); persist(); message(computer ? 'Computer campaign ready. Choose Play or Next Day.' : 'New campaign started.'); render();
    try { localStorage.setItem(COMMAND_SIDE_KEY, commandSide); }
    catch { message('Campaign started, but your command side could not be saved.'); }
  });
  $('rules').onclick = () => { pauseSimulation(); render(); $('rules-dialog').showModal(); };
  document.querySelectorAll('[data-close]').forEach(button => button.onclick = () => $(button.dataset.close).close());
  $('export').onclick = () => {
    const url = URL.createObjectURL(new Blob([JSON.stringify(state, null, 2)], { type: 'application/json' }));
    const link = document.createElement('a'); link.href = url; link.download = `winter-offensive-turn-${state.turn}.json`; link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000); message('Campaign exported.');
  };
  $('import').onclick = () => { pauseSimulation(); render(); $('save-file').click(); };
  $('save-file').onchange = async event => {
    const file = event.target.files[0]; if (!file) return;
    try {
      if (file.size > 200000) throw new Error('Save file is too large.');
      const loaded = deserialize(await file.text());
      if (busy) throw new Error('Wait for the current turn to finish.');
      pauseSimulation();
      moveHistory.length = 0;
      state = loaded;
      resetSideReports();
      if (state.ruleset === 'manual') commandSide = state.playerSide;
      if (observing()) commandSide = 'german';
      selected = state.units.find(unit => unit.side === commandSide && unit.hp > 0 && !unit.exited)?.id || null;
      const importedSelection = state.units.find(unit => unit.id === selected);
      if (importedSelection) cursor = { x: importedSelection.x, y: importedSelection.y };
      mode = state.ruleset === 'manual' && state.orderPhase === 'combat' ? 'attack' : 'move';
      persist(); message('Campaign restored.'); render();
    } catch (error) { message(`Import failed: ${error.message}`); }
    event.target.value = '';
  };
  new MutationObserver(draw).observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });
  render();
})(typeof globalThis !== 'undefined' ? globalThis : this);