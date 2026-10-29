(function (root) {
  'use strict';
  const ROT = typeof module === 'object' ? require('./rot.js') : root.ROT;
  const createManualRules = typeof module === 'object' ? require('./manual-rules.js') : root.createManualRules;
  const referenceMap = typeof module === 'object' ? require('./reference-map.js') : root.ReferenceMap;
  const aidCard = typeof module === 'object' ? require('./aid-card.js') : root.AidCard;
  let manualRules;
  const COLS = 15, ROWS = 12, LAST_TURN = 12;
  const REFERENCE_RIVER_WIDTH = 3;
  const DIRECTIONS = ROT.DIRS[6];
  const TERRAIN = {
    plain: { cost: 1, defense: 1, symbol: '.' },
    forest: { cost: 2, defense: 1.5, symbol: '*' },
    ridge: { cost: 3, defense: 1.7, symbol: '^' },
    river: { cost: 3, defense: 1.3, symbol: '~' },
    road: { cost: 1, defense: 1, symbol: ':' },
    town: { cost: 1, defense: 1.8, symbol: '#' },
    westwall: { cost: 4, defense: 2.75, symbol: '}' }
  };
  const opponent = side => side === 'german' ? 'american' : 'german';
  function unitPresentation(unit) {
    if (/\b(AIRB|FJ|Airborne)\b/i.test(unit.name)) return { symbol: 'P', label: 'Airborne / paratroopers' };
    if (/\b(PZGR|CALV)\b/i.test(unit.name)) return { symbol: 'M', label: 'Mechanised infantry / cavalry' };
    return unit.type === 'armor' ? { symbol: '=', label: 'Armour' } : { symbol: 'X', label: 'Infantry' };
  }
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
  const mapDimensions = state => state.referenceVersion === 2 ?
    { cols: referenceMap.COLS, rows: referenceMap.ROWS } : { cols: COLS, rows: ROWS };
  const riverCourses = [
    { size: 'major', points: [[7, 258], [30, 245], [53, 258], [76, 244], [99, 258], [99, 284], [122, 298], [145, 284], [168, 298], [191, 284], [191, 258], [214, 244], [237, 258], [260, 244], [283, 258], [306, 244], [329, 258], [352, 244], [375, 258], [375, 284], [398, 298], [421, 284], [444, 298], [467, 284], [490, 298], [513, 284], [513, 258], [536, 244], [559, 258], [582, 244], [605, 258], [628, 244], [651, 258], [651, 284], [674, 298], [697, 285], [720, 298], [720, 324]] },
    { size: 'minor', points: [[283, 18], [283, 44], [306, 58], [306, 84], [329, 98], [352, 84], [375, 98], [375, 124], [398, 138], [421, 124], [421, 98], [444, 84], [467, 98], [490, 84], [513, 98], [536, 84], [559, 98], [559, 124], [582, 138], [605, 124], [628, 138], [651, 124], [674, 138]] },
    { size: 'minor', points: [[7, 364], [30, 378], [53, 364], [76, 378], [99, 364], [122, 378], [122, 404], [145, 418], [168, 404], [191, 418], [214, 404], [237, 418], [260, 404], [260, 378], [283, 364], [306, 378], [329, 364], [352, 378], [375, 364], [398, 378], [421, 364], [444, 378], [444, 404], [467, 418], [490, 404], [513, 418], [513, 444], [536, 458], [559, 444], [559, 418], [582, 404], [605, 418], [628, 404], [651, 418], [674, 404], [697, 418], [720, 404], [720, 378]] }
  ];
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
  function createGame(seed = 19441216, ruleset = 'classic', playerSide = 'german', referenceVersion = 2, options = {}) {
    seed = parseSeed(seed);
    if (!['classic', 'manual'].includes(ruleset)) throw new Error('Unknown campaign ruleset.');
    if (!options || typeof options !== 'object' || Array.isArray(options) ||
      (options.playMode !== undefined && !['human','computer','hotseat'].includes(options.playMode)) ||
      (options.appleSupplement !== undefined && typeof options.appleSupplement !== 'boolean')) {
      throw new Error('Choose valid campaign options and player controls.');
    }
    if (options.playMode === 'hotseat' && (ruleset !== 'manual' || referenceVersion !== 2 || options.aidCard === false)) {
      throw new Error('Two-player hotseat requires the new Manual player-aid rules.');
    }
    if (ruleset === 'manual' && ![1, 2].includes(referenceVersion)) throw new Error('Unknown Manual map revision.');
    const state = { version: 1, turn: 1, phase: 'german', weather: 'Fog', seed, initialSeed: seed,
      map: {}, units: [], log: [], score: 0, winner: null };
    for (let row = 0; row < ROWS; row++) {
      for (let col = 0; col < COLS; col++) {
        const tile = { ...position(col, row), col, row, terrain: 'plain', owner: null };
        if ((col * 7 + row * 11) % 13 < 4) tile.terrain = 'forest';
        if (row < 4 && col > 5 && col < 10) tile.terrain = 'ridge';
        if (col === 3 || (col === 8 && row > 4)) {
          tile.terrain = 'river';
          tile.riverSize = col === 3 ? 'major' : 'minor';
        }
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
    if (ruleset === 'manual') return manualRules.initialize(state, playerSide, options.playMode !== 'computer', referenceVersion, options);
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
    if (state.pendingBattle || state.pendingDefense) return 'Complete the pending defensive battle, retreat or advance before issuing another order.';
    if (state.winner) return 'The campaign has finished.';
    if (state.playMode === 'computer') return 'Spectator mode: the computers issue orders; you can inspect either army.';
    if (!unit || unit.hp <= 0 || unit.exited) return 'Select a living formation.';
    if (unit.side !== state.phase) return 'Wait for this army’s orders phase.';
    if (order === 'move' && state.ruleset === 'manual' && state.turn === 1 && unit.side === 'german') {
      return 'German opening turn: no movement on day 1; Germany starts in combat.';
    }
    if (state.ruleset === 'manual' && state.orderPhase !== (order === 'move' ? 'movement' : 'combat')) {
      return order === 'move' ? 'Movement is closed. End combat to reach the next day.' : 'End Movement before issuing attacks.';
    }
    if (unit.fired) return 'This formation has already attacked and spent its movement today.';
    if (order === 'move' && !reachable(state, unit).size) {
      return unit.mp <= 0 ? `${unit.name} has no movement points remaining this phase.` :
        'No legal destinations are reachable with the remaining movement points. Terrain costs, occupied hexes and enemy zones may block movement.';
    }
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
  function planGermanCardMove(state, unit, options, enemies, friends, objectives, current, weak) {
    const towns = objectives.filter(tile => tile.sourcePoints);
    const germanOn = tile => {
      const occ = occupant(state, tile);
      return Boolean(occ && occ.side === 'german');
    };
    // The unit on a captured town stays there. Bastogne is 100 VP per computer turn; leaving it gives the award back.
    if (current.sourcePoints && germanOn(current)) return null;
    const needs = towns.filter(tile => !germanOn(tile));
    const exits = Object.values(state.map).filter(tile => tile.exit);
    const nearest = tiles => tiles.slice().sort((a, b) => distance(unit, a) - distance(unit, b))[0];
    const bastogne = towns.find(tile => tile.name === 'BASTOGNE');
    const townTarget = needs.slice().sort((a, b) => {
      const score = tile => distance(unit, tile) - (tile.name === 'BASTOGNE' ? 8 : 3);
      return score(a) - score(b);
    })[0];
    const exitOpen = tile => tile.exit && !occupant(state, tile) && !enemiesNear(state, tile, 'german');
    const exitTarget = nearest(exits.filter(exitOpen)) || nearest(exits);
    // Exit VP is CP × (16 − turn) / 4. Turns 1–4 stay on the towns so an opening rush cannot farm that multiplier.
    const allowExits = !weak && state.turn >= 5;
    if (allowExits && current.exit && !enemiesNear(state, current, unit.side)) return null;
    const bastogneClose = bastogne && !germanOn(bastogne) && distance(unit, bastogne) <= 3;
    let target = townTarget || nearest(towns) || exitTarget;
    if (weak) target = nearest(towns.filter(germanOn)) || nearest(towns) || target;
    // Formations still beside an unheld Bastogne keep attacking it. Everyone else, from turn 5, takes an open exit.
    else if (allowExits && exitTarget && !bastogneClose) target = exitTarget;
    else if (townTarget) target = townTarget;
    const scarceFuel = state.resourceVersion === 1 ? manualRules.provisional.scarceFuel : 35;
    const rank = tile => {
      if (!allowExits && tile.exit) return -Infinity;
      const adjacent = enemies.filter(enemy => distance(tile, enemy) === 1).length;
      const support = friends.filter(friend => distance(tile, friend) <= 2).length;
      const toward = target ? -distance(tile, target) * 3 : 0;
      const town = tile.sourcePoints && !germanOn(tile) ? tile.name === 'BASTOGNE' ? 8 : 5 : 0;
      const exitBonus = allowExits && tile.exit && target && target.exit ? 4 : 0;
      const travel = options.get(key(tile.x, tile.y))?.path.length || 0;
      return toward + town + exitBonus + (current.col - tile.col) * 0.5 - adjacent * (weak ? 6 : 0.35) +
        Math.min(3, support) * 0.4 - travel * (state.fuel < scarceFuel ? 1.1 : 0.12);
    };
    const ranked = [...options.keys()].map(tileKey => state.map[tileKey]).sort((a, b) => rank(b) - rank(a));
    const destination = ranked.slice(0, 8).find(tile => {
      if (!(rank(tile) > rank(current) + 0.2) || (!allowExits && tile.exit)) return false;
      const candidate = { ...unit, x: tile.x, y: tile.y };
      const projected = { ...state, units: state.units.map(counter => counter.id === unit.id ? candidate : counter) };
      const supply = manualRules.localSupply(projected, candidate);
      return supply !== 'isolated' || (tile.sourcePoints && !germanOn(tile));
    });
    if (!destination) return null;
    const reason = weak ? 'preserve a weakened formation and seek safer ground' :
      destination.exit ? 'continue a breakthrough toward an exit after the town offensive is underway' :
        `drive on ${target && target.name || 'Bastogne'} and the other key towns`;
    return { tile: destination, path: options.get(key(destination.x, destination.y)).path, reason };
  }
  function planMovement(state, unit) {
    const options = reachable(state, unit);
    if (!options.size) return null;
    const enemies = state.units.filter(enemy => enemy.side !== unit.side && enemy.hp > 0 && !enemy.exited);
    const friends = state.units.filter(friend => friend.side === unit.side && friend.hp > 0 && !friend.exited && friend.id !== unit.id);
    const objectives = Object.values(state.map).filter(tile => tile.value);
    const current = state.map[key(unit.x, unit.y)], weak = unit.hp < unit.maxHp * 0.35;
    if (aidCard.active(state) && unit.side === 'german') {
      return planGermanCardMove(state, unit, options, enemies, friends, objectives, current, weak);
    }
    const westernSector = state.referenceVersion === 2 ? Math.floor(mapDimensions(state).cols / 2) : 8;
    const approaching = enemies.filter(enemy => state.map[key(enemy.x, enemy.y)].col <= westernSector ||
      objectives.some(tile => tile.owner === unit.side && distance(enemy, tile) <= 4));
    if (current.value && current.owner === unit.side && !weak &&
      (unit.side === 'german' || !approaching.length || enemiesNear(state, current, unit.side))) return null;
    let targets = weak ? objectives.filter(tile => tile.owner === unit.side && !enemiesNear(state, tile, unit.side)) :
      objectives.filter(tile => tile.owner !== unit.side || !occupant(state, tile));
    if (state.referenceVersion === 2 && unit.side === 'german' && !weak) {
      targets = objectives.filter(tile => tile.owner !== unit.side);
    }
    if (!weak && unit.side === 'american' && approaching.length) targets = approaching;
    if (!targets.length) targets = Object.values(state.map).filter(tile =>
      state.referenceVersion === 2 && unit.side === 'german' && !weak ? tile.exit :
        tile.col === (unit.side === 'german' && !weak ? 0 : unit.side === 'german' ? mapDimensions(state).cols - 1 : 0));
    const target = targets.slice().sort((a, b) =>
      (distance(unit, a) - (a.value || 0) * 0.6) - (distance(unit, b) - (b.value || 0) * 0.6))[0];
    const rank = tile => {
      const adjacent = enemies.filter(enemy => distance(tile, enemy) === 1);
      const support = friends.filter(friend => distance(tile, friend) <= 2).length;
      const terrain = state.ruleset === 'manual' ? manualRules.terrainDefense(state, tile, unit.side) : TERRAIN[tile.terrain].defense;
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
          if (!move(state, unit.id, order.tile.x, order.tile.y)) {
            if (state.log[0] && state.log[0].includes('BRIDGE BLOWN')) {
              return { side, id: unit.id, text: state.log[0] };
            }
            throw new Error('Computer movement became invalid.');
          }
          const arrived = state.map[key(unit.x, unit.y)];
          const blown = state.log[0] && state.log[0].includes('BRIDGE BLOWN');
          return { side, id: unit.id, path: [from, ...order.path],
            text: blown ? state.log[0] :
              `${unit.name} moves to column ${arrived.col + 1}, row ${arrived.row + 1} to ${order.reason}.` };
        }
      } else {
        const germanCard = manual && aidCard.active(state) && side === 'german';
        if (!germanCard) progress.index++;
        if (!manual) progress.stage = 'movement';
        if (!germanCard && (unit.hp <= 0 || unit.exited || unit.fired)) continue;
        const battle = germanCard ? manualRules.nextBattle(state, side) :
          manual ? manualRules.planBattle(state, unit) : planBattle(state, unit);
        if (germanCard) {
          if (!battle) { progress.index = units.length; continue; }
          progress.index = Math.min(units.length, progress.index + 1);
        }
        if (battle) {
          const lead = battle.ids ? state.units.find(counter => counter.id === battle.ids[0]) || unit : unit;
          const before = state.units.map(counter => ({ id: counter.id, hp: counter.hp, x: counter.x, y: counter.y }));
          const from = { x: lead.x, y: lead.y }, to = { x: battle.defender.x, y: battle.defender.y };
          if (manual) manualRules.resolveAttack(state, battle.ids, battle.defender.id, battle.options);
          else if (!attack(state, lead.id, battle.defender.id)) throw new Error('Computer attack became invalid.');
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
          return { side, id: lead.id, path: [from, to], paths, battle: true,
            text: `${lead.name}${manual && battle.ids.length > 1 ? ' and supporting formations' : ''} attacks ${battle.defender.name} to ${battle.reason}. ${changes.join('. ')}.` };
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
      parsed.appleSupplementVersion !== undefined || parsed.strategyPressure !== undefined || parsed.airInterdiction !== undefined ||
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
    if (aidCard.active(state)) return aidCard.standing(state);
    const target = 8;
    const leader = state.winner || (state.score >= target ? 'german' : 'american');
    return { points: state.score, target, leader,
      status: state.winner ? `${leader.toUpperCase()} VICTORY` : `${leader.toUpperCase()} ON COURSE`,
      detail: state.winner ? 'Campaign finished.' :
        state.score >= target ? 'German target reached; hold it through the final victory check.' :
          `Germany needs ${target - state.score} more points to reach its target.`,
      provisional: !state.winner };
  }
  const api = { createGame, unitPresentation, neighbors, riverCourses, referenceMap, aidCard, mapDimensions, occupant, distance, reachable, move, attack, combatOdds,
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
  let state = createGame(freshSeed(), 'manual'), commandSide = 'german', selected = 'g1', mode = 'move', busy = false;
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
  let lastReportPhase = null;
  function showPhaseReport() {
    if (!manualRules.supplement(state) || busy) return;
    const report = manualRules.phaseReport(state);
    $('phase-report-date').textContent = `${report.date} / ${report.weather} / ${report.phase} / Score ${report.score}`;
    $('phase-report-armies').replaceChildren();
    for (const army of report.armies) {
      const row = document.createElement('tr');
      for (const value of [army.side === 'american' ? 'Allied' : 'German', army.supply, army.victory, army.artillery, army.fuel ?? 'N/A']) {
        const cell = document.createElement('td'); cell.textContent = String(value); row.append(cell);
      }
      $('phase-report-armies').append(row);
    }
    const air = report.interdiction;
    const movementReductions = Object.entries(air?.movement || {}).filter(([, loss]) => loss > 0);
    $('phase-report-interdiction').textContent = air ?
      `Turn ${air.turn} Allied air interdiction: German fuel reduced ${air.fuel} points; supply reduced ${air.supply} points.${movementReductions.length ? ' Movement reductions per formation:' : ''}` : '';
    $('phase-report-movement').replaceChildren();
    for (const [id, loss] of movementReductions) {
      const unit = state.units.find(unit => unit.id === id);
      const item = document.createElement('li'); item.textContent = `${unit.name}: ${loss} MP`; $('phase-report-movement').append(item);
    }
    if (!$('phase-report-dialog').open) $('phase-report-dialog').showModal();
    $('phase-report-title').focus({ preventScroll: true });
    $('phase-report-dialog').scrollTop = 0;
    lastReportPhase = `${state.turn}:${state.phase}:${state.orderPhase}`;
  }
  function automaticPhaseReport() {
    $('phase-report-button').hidden = !manualRules.supplement(state);
    $('phase-report-button').disabled = busy;
    if (manualRules.supplement(state) && !busy && !observing() && !state.winner &&
      lastReportPhase !== `${state.turn}:${state.phase}:${state.orderPhase}` && !document.querySelector('dialog[open]')) {
      showPhaseReport();
    }
  }
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
  const center = tile => state.referenceVersion === 2 ? referenceMap.center(tile) :
    { x: 30 + tile.x * 23, y: 31 + tile.y * 40 };
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
  function drawReferenceTerrain(ctx, terrain, point, palette) {
    ctx.fillStyle = palette[terrain];
    if (terrain === 'forest' || terrain === 'ridge') {
      const marks = terrain === 'forest' ? [[0,-8],[-6,-2],[6,-2],[0,4],[-6,10],[6,10]] : [[0,-5],[-6,5],[6,5]];
      for (const [dx, dy] of marks) ctx.fillRect(point.x + dx - 4, point.y + dy - 2, 8, 3);
    } else if (terrain === 'town') {
      ctx.fillRect(point.x - 6, point.y - 6, 12, 12);
    } else if (terrain === 'westwall') {
      ctx.strokeStyle = palette.westwall; ctx.lineWidth = 2; ctx.beginPath();
      ctx.moveTo(point.x + 2, point.y - 11);
      ctx.lineTo(point.x - 2, point.y - 4); ctx.lineTo(point.x + 2, point.y + 3);
      ctx.lineTo(point.x - 2, point.y + 11); ctx.stroke();
    }
  }
  const alternateColor = (edge, palette, ink) => edge === 'north' ? palette.highlight :
    edge === 'south' ? palette.town : edge === 'west' ? palette.forest : ink;
  function drawReferenceLegend(palette, ink) {
    for (const symbol of document.querySelectorAll('.legend-symbol')) {
      const ctx = symbol.getContext('2d'), type = symbol.dataset.symbol;
      ctx.clearRect(0, 0, 26, 26);
      if (['plain', 'river', 'north', 'west', 'south', 'german'].includes(type)) {
        ctx.beginPath();
        for (let i = 0; i < 6; i++) {
          const angle = (i * 60 - 30) * Math.PI / 180;
          const x = 13 + 11 * Math.cos(angle), y = 13 + 11 * Math.sin(angle);
          if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
        }
        ctx.closePath();
        ctx.strokeStyle = type === 'plain' || type === 'river' ? ink : alternateColor(type, palette, ink);
        ctx.lineWidth = 1; ctx.stroke();
        if (type === 'river') {
          ctx.strokeStyle = palette.river; ctx.lineWidth = REFERENCE_RIVER_WIDTH; ctx.beginPath();
          ctx.moveTo(13, 2); ctx.lineTo(13 + 11 * Math.sqrt(3) / 2, 7.5);
          ctx.lineTo(13 + 11 * Math.sqrt(3) / 2, 18.5); ctx.stroke();
        }
      } else drawReferenceTerrain(ctx, type, { x: 13, y: 13 }, palette);
    }
  }
  function drawNationality(side, x, y, palette) {
    context.save();
    context.translate(x, y);
    context.fillStyle = palette.forceWhite;
    context.beginPath();
    if (side === 'british') {
      context.font = 'bold 12px Consolas, monospace'; context.fillText('B', 0, 0);
    } else if (side === 'american') {
      for (let index = 0; index < 10; index++) {
        const angle = -Math.PI / 2 + index * Math.PI / 5;
        const radius = index % 2 ? 2.6 : 6;
        const px = Math.cos(angle) * radius, py = Math.sin(angle) * radius;
        if (index === 0) context.moveTo(px, py); else context.lineTo(px, py);
      }
      context.closePath(); context.fill();
    } else {
      context.fillRect(-3, -6, 6, 12);
      context.fillRect(-6, -3, 12, 6);
      context.fillStyle = palette.forceBlack;
      context.fillRect(-2, -5, 4, 10);
      context.fillRect(-5, -2, 10, 4);
    }
    context.restore();
  }
  function shownCp(unit) {
    return state.ruleset === 'manual' ? manualRules.displayedCombatPoints(state, unit) : unit.hp;
  }
  function shownMax(unit) {
    return state.ruleset === 'manual' ? unit.maxHp + manualRules.enemyStrengthOffset(state, unit) : unit.maxHp;
  }
  function draw() {
    const width = state.referenceVersion === 2 ? 920 : 740, height = state.referenceVersion === 2 ? 475 : 512;
    if (canvas.width !== width || canvas.height !== height) { canvas.width = width; canvas.height = height; }
    canvas.style.setProperty('--map-ratio', width / height);
    const style = getComputedStyle(document.documentElement);
    const dark = document.documentElement.dataset.theme === 'dark';
    const reference = state.referenceVersion === 2;
    const ink = style.getPropertyValue(reference ? '--cp-text' : dark ? '--cp-text' : '--cp-surface').trim();
    const paper = style.getPropertyValue(reference ? '--cp-bg-elevated' : dark ? '--cp-surface' : '--cp-text').trim();
    document.querySelector('.coordinates').textContent = reference ? 'REFERENCE MAP / 19 × 11 / WEST < EAST' : 'ARDENNES SECTOR / WEST < EAST';
    for (const entry of document.querySelectorAll('.legend [data-map]')) {
      entry.hidden = (entry.dataset.map === 'reference') !== reference;
    }
    const palette = {
      plain: ink,
      forest: style.getPropertyValue('--cp-success').trim(),
      ridge: style.getPropertyValue('--cp-warning').trim(),
      river: style.getPropertyValue('--cp-link').trim(),
      road: style.getPropertyValue('--cp-warning').trim(),
      town: style.getPropertyValue('--cp-accent').trim(),
      westwall: style.getPropertyValue('--cp-warning').trim(),
      german: style.getPropertyValue('--cp-danger').trim(),
      american: style.getPropertyValue('--cp-link').trim(),
      highlight: style.getPropertyValue('--cp-warning').trim(),
      forceUS: style.getPropertyValue('--cp-force-us').trim(),
      forceGerman: style.getPropertyValue('--cp-force-german').trim(),
      forceWhite: style.getPropertyValue('--cp-force-white').trim(),
      forceBlack: style.getPropertyValue('--cp-force-black').trim()
    };
    if (aidCard.active(state) && !state.display.color) {
      for (const field of ['forest','ridge','river','road','town','westwall','german','american','highlight']) palette[field] = ink;
    }
    if (reference) drawReferenceLegend(palette, ink);
    $('british-legend').hidden = !aidCard.active(state);
    $('scoring-legend').hidden = !aidCard.active(state);
    context.fillStyle = paper; context.fillRect(0, 0, canvas.width, canvas.height);
    context.font = '12px Consolas, monospace'; context.textAlign = 'center'; context.textBaseline = 'middle';
    const unit = state.units.find(candidate => candidate.id === selected);
    const options = mode === 'move' && !observing() ? reachable(state, unit) : new Map();
    for (const tile of Object.values(state.map)) {
      const point = center(tile);
      context.fillStyle = palette[tile.terrain];
      context.globalAlpha = tile.value ? 0.5 : tile.terrain === 'plain' ? 0.04 : tile.terrain === 'river' ? 0 : reference ? 0.08 : 0.23;
      hex(tile, 1); context.fill(); context.globalAlpha = 1;
      context.strokeStyle = ink; context.globalAlpha = 0.25; context.lineWidth = 1;
      if (!aidCard.active(state) || state.display.hexes) { hex(tile); context.stroke(); }
      context.globalAlpha = 1;
      if (options.has(key(tile.x, tile.y))) {
        context.fillStyle = paper; context.beginPath();
        context.arc(point.x, point.y + 8, 5, 0, Math.PI * 2); context.fill();
        context.fillStyle = ink; context.beginPath();
        context.arc(point.x, point.y + 8, 3, 0, Math.PI * 2); context.fill();
      }
      context.fillStyle = tile.value ? ink : palette[tile.terrain];
      context.globalAlpha = tile.value ? 1 : 0.9;
      if (reference) {
        drawReferenceTerrain(context, tile.terrain, point, palette);
      } else if (tile.terrain !== 'river') {
        context.fillText(TERRAIN[tile.terrain].symbol, point.x, point.y);
      }
      context.globalAlpha = 1;
      if (state.referenceVersion === 2 && (tile.entry || tile.exit || tile.alternate)) {
        context.font = '9px Consolas, monospace';
        context.fillStyle = alternateColor(tile.alternate, palette, ink);
        context.fillText(`${tile.exit ? 'E ' : ''}${tile.entry || ''}`, point.x, point.y - 18);
        if (tile.alternate) {
          context.globalAlpha = 0.65; hex(tile, 3); context.strokeStyle = context.fillStyle; context.stroke(); context.globalAlpha = 1;
        }
        context.font = '12px Consolas, monospace';
      }
      if (aidCard.active(state)) {
        const germanVP = aidCard.edgeAward(tile, 'german') * aidCard.multiplier(state);
        const alliedVP = aidCard.edgeAward(tile, 'american');
        if (germanVP || alliedVP) {
          context.font = '9px Consolas, monospace'; context.fillStyle = ink;
          context.fillText(`${germanVP || alliedVP}VP`, point.x, point.y + 18);
          context.font = '12px Consolas, monospace';
        }
      }
    }
    const drawnRivers = reference ? referenceMap.riverSegments.map(points => ({ size: 'minor', points })) : riverCourses;
    for (const river of drawnRivers) {
      context.strokeStyle = palette.river;
      context.lineWidth = reference ? REFERENCE_RIVER_WIDTH : river.size === 'minor' ? 1.5 : 3;
      context.lineJoin = 'round';
      context.lineCap = 'round';
      context.beginPath();
      river.points.forEach(([x, y], index) => {
        if (index === 0) context.moveTo(x, y);
        else context.lineTo(x, y);
      });
      context.stroke();
    }
    for (const counter of state.units.filter(candidate => candidate.hp > 0 && !candidate.exited)) {
      const point = center(counter);
      context.fillStyle = counter.side === 'american' ? palette.forceUS : palette.forceGerman;
      context.strokeStyle = palette.forceWhite; context.lineWidth = 1;
      context.fillRect(point.x - 16, point.y - 12, 32, 24); context.strokeRect(point.x - 16, point.y - 12, 32, 24);
      drawNationality(counter.nationality === 'british' ? 'british' : counter.side, point.x - 6, point.y - 5, palette);
      context.fillStyle = palette.forceWhite;
      context.font = 'bold 11px Consolas, monospace';
      context.fillText(unitPresentation(counter).symbol, point.x + 8, point.y - 4);
      context.font = '10px Consolas, monospace'; context.fillText(`${shownCp(counter)}/${counter.mp}`, point.x, point.y + 7);
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
    $('unit-tooltip').hidden = true;
    renderManeuvers();
    draw();
    renderSideReports();
    $('date').textContent = `${15 + state.turn} DEC 1944`;
    $('turn').textContent = `TURN ${state.turn} / ${aidCard.lastTurn(state)}`;
    $('weather').textContent = state.weather.toUpperCase();
    $('seed-status').textContent = state.initialSeed === undefined ? 'LEGACY / SEED UNKNOWN' : `SEED ${state.initialSeed}`;
    $('score').textContent = aidCard.active(state) ? `${state.score} NET VP` : `${state.score} / 8`;
    const manual = state.ruleset === 'manual';
    const card = aidCard.active(state);
    const alliedRoster = document.querySelector('.order-of-battle[data-side="american"]');
    alliedRoster.querySelector('h2').textContent = card ? 'ALLIED ORDER OF BATTLE' : 'US ORDER OF BATTLE';
    alliedRoster.setAttribute('aria-label', card ? 'Allied Order of Battle' : 'US Order of Battle');
    $('american-roster').setAttribute('aria-label', card ? 'Allied formations' : 'US formations');
    document.querySelector('.army-report[data-side="american"] h2').textContent = card ? 'ALLIED / WEST' : 'AMERICAN / WEST';
    document.querySelector('.score-army[data-side="american"] h2').textContent = card ? 'ALLIED / WEST' : 'AMERICAN / WEST';
    const unitLegendLabel = '= Armour / X Infantry / P Airborne or paratroopers / M Mechanised infantry or cavalry';
    $('unit-legend').setAttribute('aria-label', unitLegendLabel);
    document.querySelector('.army-report[data-side="german"] h2 + p').textContent = card ?
      'Mission: score casualties, marked edge presence, occupied towns and timely exits. Reach a German victory level at the final check.' :
      'Mission: concentrate westward, keep supply open and capture at least eight points. Adapted Manual rules also reward exits.';
    document.querySelector('.army-report[data-side="american"] h2 + p').textContent = card ?
      'Mission: inflict German casualties, deny town and exit awards, and gain points at marked eastern positions.' :
      'Mission: hold the western objectives, delay the offensive and deny Germany eight points at the final check.';
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
    $('campaign-score-basis').textContent = card ?
      `Cumulative player-aid VP: casualties 1 per CP; German marked south/north presence ${3 * aidCard.multiplier(state)}/${10 * aidCard.multiplier(state)}, Bastogne ${50 * aidCard.multiplier(state)}, Marche/Rochefort ${25 * aidCard.multiplier(state)} each per turn; Allied marked east presence 20; German exits CP × (16 − turn/4). Town awards require occupation, not just ownership. Ratings G${state.ratings.german}/A${state.ratings.american}.` : manual ?
      'Adapted German score: objectives + exits − Allied breakthrough. Not the original SSI scoring table.' :
      'German score: currently controlled objective points.';
    $('victory-rule-title').textContent = card ? 'PLAYER-AID VICTORY / GERMAN VP − ALLIED VP' : 'EXISTING VICTORY RULE / GERMAN SCORE';
    $('campaign-score-progress').max = standing.target;
    $('campaign-score-progress').value = Math.max(0, Math.min(standing.target, standing.points));
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
    $('ruleset-status').textContent = card ? 'MANUAL / PLAYER-AID RULES' :
      manual ? 'MANUAL-BASED ADAPTATION' : 'CLASSIC RULES';
    $('manual-command').hidden = !manual;
    if (manual) {
      $('resources').textContent = `Supply ${state.generalSupply[commandSide]} / Artillery ${state.artillery[commandSide]} / German fuel ${state.fuel}`;
      $('defense-strategy').value = state.defenseStrategy;
      $('defense-strategy').disabled = observing() || busy;
    }
    $('defensive-artillery-label').hidden = !card; $('defensive-artillery').hidden = !card;
    if (card) {
      $('defensive-artillery').max = manualRules.artilleryMaximum(state, commandSide);
      $('defensive-artillery').value = state.defensiveArtillery;
      $('defensive-artillery').disabled = observing() || busy || Boolean(state.pendingBattle);
    }
    $('end-turn').textContent = manual && state.orderPhase === 'movement' ? 'End Movement' : manual ? 'End Combat / Day' : 'End Turn';
    const unit = state.units.find(candidate => candidate.id === selected);
    $('unit-name').textContent = unit ? unit.name : 'No unit selected';
    $('unit-kind').textContent = unit ? `${unit.nationality === 'british' ? 'BRITISH / ALLIED' : unit.side.toUpperCase()} / ${unitPresentation(unit).label.toUpperCase()}` : 'GERMAN COMMAND';
    $('strength').textContent = unit ? `${shownCp(unit)} / ${shownMax(unit)}` : '-';
    $('movement').textContent = unit ? String(unit.mp) : '-';
    $('supply').textContent = unit ? manual ? manualRules.localSupply(state, unit).toUpperCase() : unit.supplied ? 'CONNECTED' : 'ISOLATED' : '-';
    $('supply').dataset.status = unit ? unit.supplied ? 'connected' : 'isolated' : '';
    $('combat').textContent = unit ? `${unit.fired ? 'COMMITTED' : 'AVAILABLE'}${manual ? ` / readiness ${unit.readiness}/7` : ''}` : '-';
    for (const command of ['move', 'attack', 'end-turn']) {
      $(command).disabled = observing() || busy || Boolean(state.winner) || Boolean(state.pendingBattle || state.pendingDefense) || (command !== 'end-turn' && (!unit || unit.hp <= 0 || unit.exited)) ||
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
    for (const side of ['american', 'german']) $(`${side}-roster`).replaceChildren();
    state.units.filter(candidate => candidate.hp > 0 && !candidate.exited).forEach(counter => {
      const button = document.createElement('button');
      button.className = 'roster-unit'; button.setAttribute('aria-pressed', String(counter.id === selected));
      button.textContent = `${counter.nationality === 'british' ? 'BRITISH ' : ''}[${unitPresentation(counter).symbol}] ${counter.name}`;
      const stats = document.createElement('span'); stats.textContent = `${shownCp(counter)} / ${counter.mp}`; button.append(stats);
      button.disabled = busy;
      button.onclick = () => {
        selected = counter.id; cursor = { x: counter.x, y: counter.y };
        $('order-of-battle-dialog').close();
        render();
        canvas.focus();
        canvas.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      };
      $(`${counter.side}-roster`).append(button);
    });
    for (const side of ['american', 'german']) {
      const roster = $(`${side}-roster`);
      if (!roster.childElementCount) roster.textContent = 'No active formations.';
    }
    const reinforcementEntries = manual ? manualRules.reinforcementSchedule(state) :
      Object.entries(REINFORCEMENTS).flatMap(([turn, entries]) => entries.map(([id, name, type]) => {
        const counter = state.units.find(candidate => candidate.id === id);
        return { name, type, side: 'american', turn: Number(turn), date: `${15 + Number(turn)} DEC`,
          status: counter ? counter.hp > 0 ? 'On map' : 'Eliminated' :
            Number(turn) <= state.turn ? 'Delayed' : 'Scheduled' };
      }));
    for (const side of ['american', 'german']) {
      const list = $(`${side}-roster-reinforcements`);
      list.replaceChildren();
      for (const entry of reinforcementEntries.filter(candidate => candidate.side === side)) {
        const row = document.createElement('li');
        row.textContent = `${entry.nationality === 'british' ? 'BRITISH / ' : ''}[${unitPresentation(entry).symbol}] ${entry.name} — TURN ${entry.turn} / ${entry.date}${['Scheduled', 'Due after combat'].includes(entry.status) ? '' : ` / ${entry.status}`}`;
        list.append(row);
      }
      if (!list.childElementCount) {
        const row = document.createElement('li');
        row.textContent = 'No scheduled reinforcements.';
        list.append(row);
      }
    }
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
    const schedulePanel = $('reinforcement-schedule');
    schedulePanel.hidden = state.ruleset !== 'manual';
    const scheduleList = $('reinforcement-schedule-list');
    scheduleList.replaceChildren();
    if (state.ruleset === 'manual') {
      const incoming = manualRules.offMapReinforcements(state);
      $('reinforcement-schedule-count').textContent = `${incoming.length} WAITING`;
      if (!incoming.length) {
        const row = document.createElement('li');
        row.textContent = 'No formations are waiting to enter.';
        scheduleList.append(row);
      }
      for (const entry of incoming) {
        const row = document.createElement('li');
        row.dataset.side = entry.side;
        row.dataset.status = entry.status;
        row.textContent = `TURN ${entry.turn} / ${entry.date} — ${entry.nationality === 'british' ? 'BRITISH / ALLIED' : entry.side.toUpperCase()} / [${unitPresentation(entry).symbol}] ${entry.name} / ${entry.strength} CP${entry.entry ? ` / entry ${entry.entry}` : ''}${entry.status === 'Delayed' ? ' / Delayed' : ''}`;
        scheduleList.append(row);
      }
    }
    $('new-game').disabled = busy; $('import').disabled = busy; $('export').disabled = busy;
    automaticPhaseReport();
  }
  function renderManeuvers() {
    $('defense-panel').hidden = !state.pendingDefense;
    if (state.pendingDefense) {
      const pending = state.pendingDefense, defender = state.units.find(unit => unit.id === pending.defenderId);
      selected = defender.id; cursor = { x: defender.x, y: defender.y };
      $('defense-description').textContent = `${pending.ids.map(id => state.units.find(unit => unit.id === id).name).join(', ')} will attack ${defender.name}. Choose your strategy and defensive artillery for this battle.`;
      $('battle-defense-strategy').value = state.defenseStrategy;
      $('battle-defense-artillery').max = Math.min(state.artillery[state.playerSide], manualRules.artilleryMaximum(state, state.playerSide));
      $('battle-defense-artillery').value = Math.min(state.defensiveArtillery,state.artillery[state.playerSide]);
    }
    const task = manualRules.maneuvers.current(state);
    $('maneuver-panel').hidden = !task;
    $('maneuver-directions').replaceChildren();
    if (!task) return;
    const unit = state.units.find(unit => unit.id === task.id);
    selected = unit.id; cursor = { x: unit.x, y: unit.y };
    const eligible = state.pendingBattle.secondCombatIds || [];
    $('maneuver-description').textContent = `${eligible.length ? `Units eligible for second combat: ${eligible.map(id => state.units.find(unit => unit.id === id).name).join(', ')}. ` : ''}${unit.name}: ${task.kind}, ${task.remaining} hex(es) remaining. ${task.kind === 'advance' && !task.path.length ? 'The first step must enter the vacated enemy hex. ' : ''}${task.kind === 'retreat' ? 'Stopping early causes additional losses.' : 'Advance is optional.'} Choose 1–6 or F to finish.`;
    const labels = ['NE','E','SE','SW','W','NW'];
    const legal = manualRules.maneuvers.choices(state);
    for (let direction = 1; direction <= 6; direction++) {
      const button = document.createElement('button');
      button.textContent = `${direction} / ${labels[direction - 1]}`;
      button.disabled = busy || !legal.some(choice => choice.direction === direction);
      button.onclick = () => doManeuver(direction);
      $('maneuver-directions').append(button);
    }
    const finish = document.createElement('button');
    finish.textContent = task.kind === 'retreat' ? 'F / Stop retreat (extra losses)' : 'F / Decline further advance';
    finish.disabled = busy; finish.onclick = () => doManeuver('finish');
    $('maneuver-directions').append(finish);
  }
  function doManeuver(direction) {
    if (busy || !state.pendingBattle) return;
    try {
      const result = manualRules.chooseManeuver(state, direction);
      moveHistory.length = 0; persist();
      mode = state.orderPhase === 'combat' ? 'attack' : 'move';
      $('maneuver-error').textContent = '';
      $('action-story').textContent = result.pending ? 'Maneuver step completed; choose the next direction.' :
        `Battle maneuvers complete. ${result.units.map(unit => `${unit.name}: ${unit.loss} CP lost, ${unit.remaining} left`).join('; ')}.`;
      render();
    } catch (error) { $('maneuver-error').textContent = error.message; }
  }
  function activate(tile) {
    if (!tile || busy || state.winner) return;
    const target = occupant(state, tile);
    if (state.pendingDefense) { message('Choose defensive strategy and artillery in the pending battle panel first.'); return; }
    if (state.pendingBattle) {
      const task = manualRules.maneuvers.current(state), unit = state.units.find(unit => unit.id === task.id);
      const choice = manualRules.maneuvers.choices(state).find(choice => choice.tile === tile);
      if (choice) doManeuver(choice.direction);
      else message(`Finish ${unit.name}'s ${task.kind} using a highlighted direction or F.`);
      return;
    }
    if (observing()) {
      if (target) selected = target.id;
      render(); return;
    }
    const unit = state.units.find(candidate => candidate.id === selected);
    if (target && target.side === commandSide) {
      selected = target.id;
      const reason = orderReason(state, target, 'move');
      message(reason || `${target.name}: ${target.mp} MP available. Choose a dotted destination.`);
      render(); return;
    }
    if (mode === 'attack' && target && unit) {
      openCombatPlan(target); return;
    } else if (mode === 'move') {
      const previous = { save: JSON.stringify(state), unitId: selected };
      const option = reachable(state, unit).get(key(tile.x, tile.y));
      const from = unit ? { x: unit.x, y: unit.y } : null;
      if (move(state, selected, tile.x, tile.y)) {
        const blown = state.log[0] && state.log[0].includes('BRIDGE BLOWN');
        const arrived = state.map[key(unit.x, unit.y)];
        setActionVisual({ side: unit.side, path: [from, ...option.path] });
        $('action-story').textContent = blown ? state.log[0] :
          `${unit.name} moved along the highlighted route to column ${arrived.col + 1}, row ${arrived.row + 1}. ${unit.mp} movement points remain.`;
        moveHistory.push(previous); message(blown ? state.log[0] : 'Movement complete.'); persist();
      } else {
        if (manualRules.supplement(state)) persist();
        const blown = state.log[0] && state.log[0].includes('BRIDGE BLOWN') ? state.log[0] : '';
        message(blown || (state.ruleset === 'manual' && manualRules.movementWarning(state, unit, tile)) ||
          orderReason(state, unit, 'move') || 'Destination blocked: choose an empty dotted hex within the movement budget. Enemy zones and terrain restrict the route.');
      }
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
  canvas.addEventListener('pointermove', event => {
    const tooltip = $('unit-tooltip');
    const bounds = canvas.getBoundingClientRect();
    const x = (event.clientX - bounds.left) * canvas.width / bounds.width;
    const y = (event.clientY - bounds.top) * canvas.height / bounds.height;
    const unit = state.units.find(candidate => {
      const point = center(candidate);
      return candidate.hp > 0 && !candidate.exited &&
        Math.abs(point.x - x) <= 16 && Math.abs(point.y - y) <= 12;
    });
    if (!unit || event.pointerType === 'touch') { tooltip.hidden = true; return; }
    const manual = state.ruleset === 'manual';
    const title = document.createElement('h3');
    title.textContent = unit.name;
    const tile = state.map[key(unit.x, unit.y)];
    const details = [
      `${unit.nationality === 'british' ? 'British / Allied' : unit.side === 'american' ? 'US' : 'German'} ${unitPresentation(unit).label}`,
      `${manual ? 'Combat points' : 'Strength'}: ${shownCp(unit)} / ${shownMax(unit)}`,
      `Movement: ${unit.mp} / ${unit.mobility} MP`,
      `Supply: ${manual ? manualRules.localSupply(state, unit) : unit.supplied ? 'connected' : 'isolated'}`,
      `Combat: ${unit.fired ? 'already attacked' : 'not yet attacked'}`,
      `Terrain: ${tile.terrain} / column ${tile.col + 1}, row ${tile.row + 1}`
    ];
    if (!observing() && unit.side === commandSide) {
      const movementReason = orderReason(state, unit, 'move');
      if (movementReason) details.push(`Cannot move: ${movementReason}`);
    }
    if (manual) details.push(`Readiness: ${unit.readiness} / 7`);
    tooltip.replaceChildren(title, ...details.map(text => {
      const paragraph = document.createElement('p');
      paragraph.textContent = text;
      return paragraph;
    }));
    tooltip.hidden = false;
    const box = tooltip.getBoundingClientRect();
    const left = event.clientX + 16 + box.width > document.documentElement.clientWidth ?
      event.clientX - box.width - 16 : event.clientX + 16;
    const top = event.clientY + 16 + box.height > document.documentElement.clientHeight ?
      event.clientY - box.height - 16 : event.clientY + 16;
    tooltip.style.left = `${Math.max(8, left)}px`;
    tooltip.style.top = `${Math.max(8, top)}px`;
  });
  canvas.addEventListener('pointerleave', () => { $('unit-tooltip').hidden = true; });
  const legendTooltip = $('legend-tooltip');
  for (const icon of document.querySelectorAll('.legend-item')) {
    icon.removeAttribute('title');
    icon.setAttribute('aria-describedby', 'legend-tooltip');
    const showLegendTooltip = () => {
      legendTooltip.textContent = icon.getAttribute('aria-label');
      legendTooltip.hidden = false;
      const bounds = icon.getBoundingClientRect();
      const box = legendTooltip.getBoundingClientRect();
      legendTooltip.style.left = `${Math.max(8, Math.min(bounds.left,
        document.documentElement.clientWidth - box.width - 8))}px`;
      legendTooltip.style.top = `${Math.max(8, bounds.bottom + box.height + 8 >
        document.documentElement.clientHeight ? bounds.top - box.height - 8 : bounds.bottom + 8)}px`;
    };
    icon.addEventListener('pointerenter', showLegendTooltip);
    icon.addEventListener('focus', showLegendTooltip);
    icon.addEventListener('pointerleave', () => { legendTooltip.hidden = true; });
    icon.addEventListener('blur', () => { legendTooltip.hidden = true; });
  }
  document.addEventListener('keydown', event => {
    if ($('phase-report-dialog').open) {
      event.preventDefault(); event.stopImmediatePropagation(); $('phase-report-dialog').close(); return;
    }
    if (manualRules.supplement(state) && !busy && event.key.toLowerCase() === 'r' &&
      !event.ctrlKey && !event.altKey && !event.metaKey &&
      (event.target.closest('#combat-plan') ||
        (!document.querySelector('dialog[open]') && !['INPUT','SELECT','TEXTAREA'].includes(event.target.tagName)))) {
      event.preventDefault(); showPhaseReport(); return;
    }
    if (event.key === 'Escape') {
      $('unit-tooltip').hidden = true;
      legendTooltip.hidden = true;
      if ($('order-of-battle-dialog').open) $('order-of-battle-dialog').close();
    }
  });
  document.addEventListener('scroll', () => {
    $('unit-tooltip').hidden = true;
    legendTooltip.hidden = true;
  }, true);
  canvas.addEventListener('keydown', event => {
    if (aidCard.active(state) && !observing() && !busy && !state.winner) {
      if (state.pendingBattle && (/^[1-6]$/.test(event.key) || event.key.toLowerCase() === 'f')) {
        event.preventDefault(); doManeuver(event.key.toLowerCase() === 'f' ? 'finish' : Number(event.key)); return;
      }
      if (!state.pendingBattle && event.key.toLowerCase() === 'e') {
        event.preventDefault(); $('end-turn').click(); return;
      }
      if (state.orderPhase === 'movement' && /^[1-6]$/.test(event.key)) {
        event.preventDefault();
        const unit = state.units.find(unit => unit.id === selected);
        if (unit && unit.side === commandSide) {
          const [dx,dy] = aidCard.directions[Number(event.key) - 1];
          const tile = state.map[key(unit.x + dx, unit.y + dy)];
          if (tile) { mode = 'move'; activate(tile); }
          else message('That direction leaves the map.');
        }
        return;
      }
      if (state.orderPhase === 'movement' && event.key.toLowerCase() === 'f') {
        event.preventDefault();
        if (manualRules.finishMovement(state, selected)) {
          persist(); render(); message('Formation movement finished.');
        } else message('Choose an active friendly formation with movement points remaining.');
        return;
      }
    }
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
    $('plan-target').textContent = `Attack ${defender.name} / ${shownCp(defender)} ${manual ? 'CP' : 'strength'}`;
    $('manual-plan-options').hidden = !manual;
    $('hotseat-defense-options').hidden = state.playMode !== 'hotseat';
    if (state.playMode === 'hotseat') {
      $('plan-defense-artillery').max = state.artillery[defender.side];
      $('plan-defense-artillery').value = '0';
    }
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
        strategy: $('plan-strategy').value, artillery: Number($('plan-artillery').value),
        ...(state.playMode === 'hotseat' ? { defenseStrategy: $('plan-defense-strategy').value,
          defensiveArtillery: Number($('plan-defense-artillery').value) } : {})
      });
      const heading = document.createElement('strong'); heading.textContent = `${preview.odds.toFixed(2)}:1 nominal attack odds`;
      const note = document.createElement('p'); note.textContent = 'Actual losses and outcomes may vary.';
      const more = document.createElement('details');
      const summary = document.createElement('summary'); summary.textContent = 'Show combat details';
      const detail = document.createElement('p'); detail.textContent = preview.detail;
      const risk = document.createElement('p'); risk.textContent = preview.risk;
      more.append(summary, detail, risk);
      $('plan-preview').replaceChildren(heading, note, more);
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
        strategy: $('plan-strategy').value, artillery: Number($('plan-artillery').value),
        ...(state.playMode === 'hotseat' ? { defenseStrategy: $('plan-defense-strategy').value,
          defensiveArtillery: Number($('plan-defense-artillery').value) } : {})
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
  $('defensive-artillery').onchange = event => {
    if (!aidCard.active(state) || busy || observing() || state.pendingBattle) return;
    const value = Number(event.target.value);
    const maximum = manualRules.artilleryMaximum(state, commandSide);
    if (!event.target.validity.valid || !Number.isInteger(value) || value < 0 || value > maximum) {
      event.target.value = state.defensiveArtillery; message(`Defensive artillery must be a whole number from 0 to ${maximum}.`); return;
    }
    state.defensiveArtillery = value; persist();
  };
  $('resolve-defense').onclick = () => {
    if (!state.pendingDefense || busy) return;
    try {
      if (!$('battle-defense-artillery').validity.valid) throw new Error('Choose an available whole-number artillery commitment.');
      const result = manualRules.chooseDefense(state, $('battle-defense-strategy').value, Number($('battle-defense-artillery').value));
      $('defense-error').textContent = '';
      $('action-story').textContent = `Defensive battle resolved: ${result.units.map(unit => `${unit.name} lost ${unit.loss} CP (${unit.remaining} left)`).join('; ')}. ${result.pending ? 'Complete the pending retreat or advance.' : 'Review side reports for maneuvers.'}`;
      persist(); mode = state.orderPhase === 'combat' ? 'attack' : 'move'; render();
    } catch (error) { $('defense-error').textContent = error.message; }
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
        if (state.playMode === 'hotseat') commandSide = state.playerSide;
        if (state.ruleset === 'manual') mode = state.orderPhase === 'combat' ? 'attack' : 'move';
        if (!state.units.some(unit => unit.id === selected && unit.side === commandSide && unit.hp > 0 && !unit.exited)) {
          selected = state.units.find(unit => unit.side === commandSide && unit.hp > 0 && !unit.exited)?.id || null;
          const unit = state.units.find(unit => unit.id === selected);
          if (unit) cursor = { x: unit.x, y: unit.y };
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
    const hotseat = button.dataset.commandSide === 'hotseat';
    const newSide = computer || hotseat ? 'german' : button.dataset.commandSide;
    try {
      if (!$('german-rating').validity.valid || !$('allied-rating').validity.valid) {
        throw new Error('Ratings must be whole numbers from 1 to 9.');
      }
      state = createGame(seed, 'manual', newSide, 2, {
        scenario: $('campaign-scenario').value,
        ratings: { german: Number($('german-rating').value), american: Number($('allied-rating').value) },
        display: { color: $('campaign-color').checked, hexes: $('campaign-hexes').checked },
        ...(computer || hotseat ? { playMode: computer ? 'computer' : 'hotseat' } : {})
      });
    } catch (error) { $('restart-error').textContent = error.message; return; }
    commandSide = newSide; moveHistory.length = 0; lastReportPhase = null;
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
  $('order-of-battle-button').onclick = () => { pauseSimulation(); render(); $('order-of-battle-dialog').showModal(); };
  $('phase-report-button').onclick = showPhaseReport;
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
      lastReportPhase = null;
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