(function (root) {
  'use strict';
  function createManualRules(base) {
    const COST = { plain: 2, road: 1, town: 3, forest: 6, ridge: 4, river: 8 };
    const LEGACY_COST = { plain: 2, road: 1, town: 2, forest: 6, ridge: 8, river: 8 };
    // Strategy names are sourced; commitment factors and hidden combat formulas remain adaptations.
    const ATTACK = { recon: 0.35, probe: 0.6, assault: 1, major: 1.5 };
    const DEFENSE = { counter: 1.6, hold: 1.3, flexible: 1, delay: 0.7, withdrawal: 0.5 };
    const strategyLabel = value => ({ recon: 'Reconnaissance', probe: 'Light', assault: 'Medium', major: 'Major',
      counter: 'Counter-attack', hold: 'Hold position', flexible: 'Flexible defense (legacy)',
      delay: 'Delaying action', withdrawal: 'Withdrawal' }[value] || value);
    const LEGACY_SCHEDULE = [
      { id: 'a8', name: '101 Airborne', side: 'american', type: 'infantry', turn: 3, col: 5, row: 8, hp: 100 },
      { id: 'a9', name: '82 Airborne', side: 'american', type: 'infantry', turn: 5, col: 4, row: 3, hp: 100 },
      { id: 'a10', name: '4 Armored', side: 'american', type: 'armor', turn: 7, col: 2, row: 10, hp: 100 },
      { id: 'a11', name: '3 Armored', side: 'american', type: 'armor', turn: 7, col: 1, row: 2, hp: 100 },
      { id: 'g8', name: '116 Panzer', side: 'german', type: 'armor', turn: 3, col: 14, row: 6, hp: 100 },
      { id: 'g9', name: '9 SS Panzer', side: 'german', type: 'armor', turn: 5, col: 14, row: 3, hp: 100 }
    ];
    // Entry numbers refer to another map, so coordinates here are explicitly schematic.
    const SCHEDULE = [
      { id: 'a8', name: '101 Airborne', side: 'american', type: 'infantry', turn: 1, col: 5, row: 8, hp: 70 },
      { id: 'a9', name: '82 Airborne', side: 'american', type: 'infantry', turn: 1, col: 4, row: 3, hp: 70 },
      { id: 'a10', name: '4 Armored', side: 'american', type: 'armor', turn: 5, col: 2, row: 10, hp: 100 },
      { id: 'a11', name: '3 Armored', side: 'american', type: 'armor', turn: 5, col: 1, row: 2, hp: 100 },
      { id: 'g8', name: '116 Panzer', side: 'german', type: 'armor', turn: 1, col: 14, row: 6, hp: 100 },
      { id: 'g9', name: '9 SS Panzer', side: 'german', type: 'armor', turn: 3, col: 14, row: 3, hp: 115 },
      { id: 'g10', name: '2 SS Panzer', side: 'german', type: 'armor', turn: 3, col: 14, row: 7, hp: 115 }
    ];
    const schedule = state => state.referenceVersion === 1 ? SCHEDULE : LEGACY_SCHEDULE;
    const costs = state => state.referenceVersion === 1 ? COST : LEGACY_COST;
    function terrainDefense(state, tile) {
      return state.referenceVersion === 1 ?
        ({ forest: 2, ridge: 1.5, town: 2.25, plain: 1, road: 1, river: 1 }[tile.terrain]) :
        base.TERRAIN[tile.terrain].defense;
    }
    const otherSide = side => side === 'german' ? 'american' : 'german';
    const living = unit => unit.hp > 0 && !unit.exited;
    const log = (state, text) => { state.log.unshift(text); state.log = state.log.slice(0, 60); };
    function random(state) {
      state.seed = (Math.imul(1664525, state.seed) + 1013904223) >>> 0;
      return state.seed / 4294967296;
    }
    function localSupply(state, unit) {
      const around = [[-1, -1], [1, -1], [2, 0], [1, 1], [-1, 1], [-2, 0]]
        .map(([dx, dy]) => base.occupant(state, { x: unit.x + dx, y: unit.y + dy }));
      const open = around.map(counter => !counter || counter.side === unit.side);
      const gap = open.some((_, i) => open[i] && open[(i + 1) % 6] && open[(i + 2) % 6]);
      const friends = around.filter(counter => counter && counter.side === unit.side).length;
      const status = gap || friends >= 2 ? 'supplied' : friends === 1 ? 'limited' : 'isolated';
      return unit.side === 'american' && unit.airdropTurn === state.turn ?
        status === 'isolated' ? 'limited' : 'supplied' : status;
    }
    const modifier = status => status === 'supplied' ? 1 : status === 'limited' ? 0.75 : 0.5;
    const enemyNear = (state, tile, side) => base.neighbors(state, tile).some(next => {
      const counter = base.occupant(state, next);
      return counter && counter.side !== side;
    });
    function refreshSupply(state) {
      for (const unit of state.units.filter(living)) {
        unit.supplyState = localSupply(state, unit);
        unit.supplied = unit.supplyState === 'supplied';
      }
    }
    function beginSide(state, side) {
      state.phase = side; state.orderPhase = 'movement';
      refreshSupply(state);
      state.units.forEach(unit => { unit.defended = false; });
      for (const unit of state.units.filter(unit => living(unit) && unit.side === side)) {
        unit.fired = false;
        unit.readiness = Math.min(7, unit.readiness + 1);
        if (!enemyNear(state, unit, side)) unit.hp = Math.min(unit.maxHp, unit.hp + 1 + Math.floor(random(state) * 3));
        let allowance = unit.mobility * modifier(unit.supplyState);
        if (state.turn === 1) allowance = side === 'german' ? 0 : allowance / 2;
        if (side === 'german') {
          const air = { Overcast: 0, Cloudy: 1, Scattered: 3, Clear: 5 }[state.weather];
          allowance -= Math.floor(random(state) * (air + 1));
          if (unit.type === 'armor' && random(state) > state.fuel / 120) allowance /= 2;
        }
        unit.mp = Math.max(0, Math.floor(allowance));
      }
      if (state.turn === 1 && side === 'german') {
        state.orderPhase = 'combat';
        log(state, '16 DEC / German opening combat: no movement on turn 1.');
      }
    }
    function initialize(state, playerSide, opening = true) {
      if (!['german', 'american'].includes(playerSide)) throw new Error('Choose a valid command side.');
      Object.assign(state, { version: 2, ruleset: 'manual', playerSide, orderPhase: 'movement', weather: 'Overcast',
        generalSupply: { german: 120, american: 120 }, artillery: { german: 6, american: 4 },
        fuel: 120, exitPoints: 0, alliedBreakthrough: 0, defenseStrategy: 'flexible',
        arrivals: [], pending: [], referenceVersion: 1 });
      const strengths = { g1: 115, g2: 115, g3: 100, g4: 100, g5: 65, g6: 50, g7: 55,
        a1: 75, a2: 55, a3: 70, a4: 120, a5: 115, a6: 75, a7: 60 };
      for (const unit of state.units) {
        unit.hp = strengths[unit.id]; unit.maxHp = unit.hp;
        unit.mobility = unit.type === 'armor' ? 24 : 20;
        Object.assign(unit, { readiness: 6, defended: false, exited: false, supplyState: 'supplied' });
      }
      // The schematic starting line is brought into contact for opening combat.
      for (const [id, col, row] of [['g1', 11, 2], ['g2', 11, 4], ['g5', 11, 8]]) {
        const unit = state.units.find(unit => unit.id === id);
        unit.x = col * 2 + row % 2; unit.y = row;
      }
      state.defenseStrategy = 'hold';
      reinforce(state);
      state.dayStart = base.daySnapshot(state);
      beginSide(state, 'german');
      if (opening && playerSide === 'american') {
        computerTurn(state, 'german');
        beginSide(state, 'american');
      }
      return state;
    }
    function movementCost(state, from, tile, unit) {
      const enters = enemyNear(state, tile, unit.side);
      return costs(state)[tile.terrain] + (enters ? 4 + (enemyNear(state, from, unit.side) ? 4 : 0) : 0);
    }
    function reachable(state, unit) {
      if (!unit || !living(unit) || state.winner || state.orderPhase !== 'movement' || state.phase !== unit.side) return new Map();
      const start = base.key(unit.x, unit.y);
      const best = new Map([[start, { remaining: unit.mp, path: [] }]]);
      const pending = [state.map[start]];
      while (pending.length) {
        const current = pending.shift(), entry = best.get(base.key(current.x, current.y));
        if (entry.path.length && enemyNear(state, current, unit.side) && !(unit.side === 'german' && unit.type === 'armor')) continue;
        for (const tile of base.neighbors(state, current)) {
          if (base.occupant(state, tile)) continue;
          // Original directions 4/5/6 correspond to the westward half of this hex layout.
          if (unit.side === 'american' && unit.type === 'infantry' && enemyNear(state, current, unit.side) &&
            enemyNear(state, tile, unit.side) && tile.x >= current.x) continue;
          const remaining = entry.remaining - movementCost(state, current, tile, unit), tileKey = base.key(tile.x, tile.y);
          if (remaining < 0 || (best.has(tileKey) && best.get(tileKey).remaining >= remaining)) continue;
          best.set(tileKey, { remaining, path: [...entry.path, tile] }); pending.push(tile);
        }
      }
      best.delete(start);
      return best;
    }
    function capture(state, unit) {
      const tile = state.map[base.key(unit.x, unit.y)];
      if (tile.value && tile.owner !== unit.side) {
        tile.owner = unit.side; log(state, `${unit.name} captures ${tile.name}.`);
        base.reportOperation(state, unit.side, `${unit.name} captured ${tile.name}, worth ${tile.value} objective points.`);
        base.reportOperation(state, otherSide(unit.side), `${tile.name} was lost to ${unit.name}.`);
      }
    }
    function score(state) {
      state.score = Math.max(0, Object.values(state.map).reduce((sum, tile) =>
        sum + (tile.owner === 'german' ? tile.value || 0 : 0), 0) + state.exitPoints - state.alliedBreakthrough);
    }
    function move(state, id, x, y) {
      const unit = state.units.find(unit => unit.id === id), option = reachable(state, unit).get(base.key(x, y));
      if (!option) return false;
      for (const tile of option.path) {
        unit.x = tile.x; unit.y = tile.y; unit.readiness = Math.max(0, unit.readiness - 2);
        capture(state, unit);
        if (unit.side === 'german' && unit.type === 'armor') state.fuel = Math.max(0, state.fuel - Math.ceil(costs(state)[tile.terrain] / 2));
      }
      unit.mp = option.remaining;
      if (enemyNear(state, unit, unit.side) && !(unit.side === 'german' && unit.type === 'armor')) unit.mp = 0;
      refreshSupply(state); score(state);
      const tile = state.map[base.key(unit.x, unit.y)];
      base.reportOperation(state, unit.side, `${unit.name} moved to column ${tile.col + 1}, row ${tile.row + 1}. ${unit.mp} movement points remain; readiness is ${unit.readiness}/7.`);
      return true;
    }
    function strength(state, unit, defending = false) {
      const terrain = defending ? terrainDefense(state, state.map[base.key(unit.x, unit.y)]) : 1;
      return unit.hp * (1 + unit.readiness / 10) * modifier(localSupply(state, unit)) * terrain;
    }
    const combatOdds = (state, attacker, defender) => strength(state, attacker) / Math.max(1, strength(state, defender, true));
    function eligibleAttackers(state, defender) {
      if (!defender || !living(defender) || defender.side === state.phase || defender.defended ||
        state.orderPhase !== 'combat' || state.winner) return [];
      return state.units.filter(unit => living(unit) && unit.side === state.phase && !unit.fired &&
        base.distance(unit, defender) === 1 && localSupply(state, unit) !== 'isolated');
    }
    function lose(state, unit, amount) {
      unit.hp = Math.max(0, unit.hp - amount);
      if (unit.hp > 0 && unit.hp < 20 && amount > 0 && random(state) < 0.25) unit.hp = 0;
      if (unit.hp === 0) log(state, `${unit.name} eliminated.`);
    }
    function retreat(state, unit, count, attackers) {
      const path = [];
      const visited = new Set([base.key(unit.x, unit.y)]);
      if (localSupply(state, unit) !== 'isolated') {
        while (living(unit) && path.length < count) {
          const next = base.neighbors(state, unit).filter(tile => !base.occupant(state, tile) && !visited.has(base.key(tile.x, tile.y)))
            .sort((a, b) => {
              const rank = tile => Math.min(...attackers.map(attacker => base.distance(tile, attacker))) * 10 +
                (unit.side === 'german' ? tile.col : -tile.col);
              return rank(b) - rank(a);
            })[0];
          if (!next) break;
          unit.x = next.x; unit.y = next.y; unit.readiness = Math.max(0, unit.readiness - 2);
          visited.add(base.key(next.x, next.y)); path.push(base.key(next.x, next.y)); capture(state, unit);
          if (enemyNear(state, unit, unit.side)) lose(state, unit, 1 + Math.floor(random(state) * 6));
        }
      }
      if (living(unit) && path.length < count) lose(state, unit, (count - path.length) * (2 + Math.floor(random(state) * 4)));
      return path;
    }
    function resolveAttack(state, ids, defenderId, options = {}) {
      const defender = state.units.find(unit => unit.id === defenderId), eligible = eligibleAttackers(state, defender);
      if (!Array.isArray(ids) || !ids.length || ids.length > 6 || new Set(ids).size !== ids.length ||
        ids.some(id => !eligible.some(unit => unit.id === id))) throw new Error('Choose 1-6 eligible adjacent, non-isolated attackers in the combat phase.');
      let attackStrategy = options.strategy || 'assault';
      let defenseStrategy = options.defenseStrategy || (defender.side === state.playerSide ? state.defenseStrategy : 'flexible');
      const artillery = options.artillery === undefined ? 0 : options.artillery;
      if (!Object.hasOwn(ATTACK, attackStrategy) || !Object.hasOwn(DEFENSE, defenseStrategy) ||
        !Number.isInteger(artillery) || artillery < 0 || artillery > Math.min(ids.length, state.artillery[state.phase])) {
        throw new Error('Invalid combat strategy or artillery allocation.');
      }
      const attackers = ids.map(id => state.units.find(unit => unit.id === id)), side = state.phase, defenseSide = defender.side;
      const before = [...attackers, defender].map(unit => ({ id: unit.id, name: unit.name, side: unit.side, hp: unit.hp }));
      const original = { x: defender.x, y: defender.y }, events = [];
      const attackerPositions = attackers.map(unit => ({ id: unit.id, x: unit.x, y: unit.y }));
      if (random(state) < (side === 'german' ? 0.04 : 0.08) * ATTACK[attackStrategy]) {
        const strategies = state.referenceVersion === 1 ? Object.keys(ATTACK) : ['probe', 'assault', 'major'];
        attackStrategy = strategies[Math.floor(random(state) * strategies.length)]; events.push('Attacking orders altered.');
      }
      if (random(state) < 0.06 * DEFENSE[defenseStrategy]) {
        const strategies = state.referenceVersion === 1 ? ['counter', 'hold', 'delay', 'withdrawal'] : ['hold', 'flexible', 'delay'];
        defenseStrategy = strategies[Math.floor(random(state) * strategies.length)]; events.push('Defending orders altered.');
      }
      const defensiveArtillery = Math.min(2, state.artillery[defenseSide]);
      state.artillery[side] -= artillery; state.artillery[defenseSide] -= defensiveArtillery;
      function suppliedStrength(unit, defending) {
        const delivered = random(state) < Math.min(1, state.generalSupply[unit.side] / 120);
        if (!delivered) events.push(`${unit.name}: general supply failed; strength halved.`);
        return strength(state, unit, defending) * (delivered ? 1 : 0.5);
      }
      const attackTotal = attackers.reduce((sum, unit) => sum + suppliedStrength(unit, false), 0) + artillery * 10;
      const defenseTotal = suppliedStrength(defender, true) + defensiveArtillery * 10;
      const odds = attackTotal / Math.max(1, defenseTotal);
      const intensity = ATTACK[attackStrategy] + DEFENSE[defenseStrategy];
      const roll = random(state);
      const defenseLoss = Math.max(1, Math.round((odds * 5 + roll * 8) * intensity));
      const attackLoss = Math.max(0, Math.round((8 / Math.max(0.25, odds) + (1 - roll) * 4) * intensity));
      const totalHp = attackers.reduce((sum, unit) => sum + unit.hp, 0);
      attackers.forEach(unit => {
        lose(state, unit, Math.round(attackLoss * unit.hp / totalHp)); unit.fired = true; unit.mp = 0;
        if (unit.side === 'german' && unit.type === 'armor') state.fuel = Math.max(0, state.fuel - 2);
      });
      lose(state, defender, defenseLoss); defender.defended = true;
      const airChance = { Overcast: 0, Cloudy: 0.1, Scattered: 0.3, Clear: 0.5 }[state.weather];
      if (random(state) < airChance) {
        [...attackers, defender].filter(unit => living(unit) && unit.side === 'german')
          .forEach(unit => lose(state, unit, 1 + Math.floor(random(state) * 6)));
        events.push('Allied air support inflicted additional German losses.');
      }
      for (const faction of [side, defenseSide]) {
        const committed = faction === side ? attackTotal : defenseTotal;
        state.generalSupply[faction] = Math.max(0, state.generalSupply[faction] - Math.ceil(committed * intensity / 20));
      }
      const retreatCount = odds + roll > 1.7 ? (defenseStrategy === 'delay' ? 3 : odds > 2 ? 2 : 1) : 0;
      const retreatPath = living(defender) ? retreat(state, defender, retreatCount, attackers) : [];
      if (retreatCount) events.push(`${defender.name}: retreat ${retreatPath.length}/${retreatCount} hexes; unfulfilled retreat causes losses.`);
      const attackerRetreats = [];
      if (!retreatCount && living(defender) && odds + roll < 0.65) {
        for (const unit of attackers.filter(living)) {
          const path = retreat(state, unit, 1, [defender]);
          attackerRetreats.push({ id: unit.id, path });
          events.push(`${unit.name}: retreat ${path.length}/1 hex; unfulfilled retreat causes losses.`);
        }
      }
      const advanceLimit = retreatCount + (defender.hp === 0 ? 1 : 0);
      const advancePaths = [];
      if (!base.occupant(state, original)) {
        const advancer = attackers.find(living);
        if (advancer) {
          let destination = state.map[base.key(original.x, original.y)];
          const path = [];
          while (destination && path.length < advanceLimit) {
            advancer.x = destination.x; advancer.y = destination.y; advancer.readiness = Math.max(0, advancer.readiness - 2);
            capture(state, advancer); path.push(base.key(destination.x, destination.y));
            if (enemyNear(state, advancer, advancer.side)) break;
            destination = base.neighbors(state, advancer).find(tile => !base.occupant(state, tile) &&
              !path.includes(base.key(tile.x, tile.y)) && (advancer.side === 'german' ? tile.col < advancer.x / 2 : tile.col > advancer.x / 2));
          }
          advancePaths.push({ id: advancer.id, path });
          events.push(`${advancer.name} advances ${path.length} hex(es).`);
        }
      }
      if (attackerRetreats.length && living(defender)) {
        const vacated = attackerPositions.find(position => !base.occupant(state, position));
        if (vacated && base.distance(defender, vacated) === 1) {
          defender.x = vacated.x; defender.y = vacated.y; defender.readiness = Math.max(0, defender.readiness - 2);
          capture(state, defender);
          advancePaths.push({ id: defender.id, path: [base.key(defender.x, defender.y)] });
          events.push(`${defender.name} advances into the attackers' vacated hex.`);
        }
      }
      if (advanceLimit >= 3) {
        attackers.filter(living).forEach(unit => {
          if (random(state) < 0.5) { unit.fired = false; events.push(`${unit.name} may attack again.`); }
        });
      }
      refreshSupply(state); score(state);
      const after = before.map(snapshot => {
        const unit = state.units.find(unit => unit.id === snapshot.id);
        return { ...snapshot, remaining: unit.hp, loss: snapshot.hp - unit.hp };
      });
      log(state, `${attackers.map(unit => unit.name).join(' + ')} > ${defender.name}: ${odds.toFixed(1)}:1 / ${attackStrategy} vs ${defenseStrategy}.`);
      events.forEach(event => log(state, event));
      for (const unit of after) {
        const retreatResult = unit.id === defender.id ? retreatPath : (attackerRetreats.find(entry => entry.id === unit.id)?.path || []);
        const advance = advancePaths.find(entry => entry.id === unit.id);
        base.reportOperation(state, unit.side, `${unit.name} ${unit.id === defender.id ? `defended against ${attackers.map(unit => unit.name).join(' and ')}` : `attacked ${defender.name}`} at ${odds.toFixed(1)}:1 odds (${strategyLabel(attackStrategy)} attack / ${strategyLabel(defenseStrategy)} defense). It lost ${unit.loss} combat points, leaving ${unit.remaining}. ${unit.remaining === 0 ? 'Our formation was eliminated.' : retreatResult.length ? `It retreated ${retreatResult.length} hex(es).` : advance ? `It advanced ${advance.path.length} hex(es).` : 'It remained in position.'}`);
      }
      for (const event of events.filter(event => event.includes('supply') || event.includes('air support') || event.includes('again') || event.includes('altered'))) {
        base.reportOperation(state, side, event);
        base.reportOperation(state, defenseSide, event);
      }
      return { odds, attackStrategy, defenseStrategy, artillery, defensiveArtillery, units: after, events, retreatPath, attackerRetreats, advancePaths };
    }
    function computerTurn(state, side) {
      if (state.winner) return;
      if (state.phase !== side) beginSide(state, side);
      for (const unit of state.units.filter(unit => living(unit) && unit.side === side)) {
        const order = base.planMovement(state, unit);
        if (order) move(state, unit.id, order.tile.x, order.tile.y);
      }
      state.orderPhase = 'combat';
      for (const unit of state.units.filter(unit => living(unit) && unit.side === side)) {
        const battle = planBattle(state, unit);
        if (battle) resolveAttack(state, battle.ids, battle.defender.id, battle.options);
      }
    }
    function planBattle(state, unit) {
      if (!living(unit) || unit.fired || unit.hp < unit.maxHp * 0.3) return null;
      const targets = state.units.filter(defender => living(defender) && defender.side !== unit.side &&
        eligibleAttackers(state, defender).includes(unit)).map(defender => {
        const attackers = eligibleAttackers(state, defender).filter(attacker => attacker.hp >= attacker.maxHp * 0.3);
        const artillery = Math.min(attackers.length, state.artillery[unit.side], state.generalSupply[unit.side] < 40 ? 1 : 3);
        const ids = attackers.map(attacker => attacker.id);
        const defenseStrategy = state.playMode !== 'computer' && defender.side === state.playerSide ? state.defenseStrategy :
          state.map[base.key(defender.x, defender.y)].value ? 'hold' : defender.hp < defender.maxHp * 0.4 ? 'delay' : 'flexible';
        const options = { strategy: state.generalSupply[unit.side] < 35 ? 'probe' : 'assault', artillery, defenseStrategy };
        const odds = preview(state, ids, defender.id, options).odds;
        return { defender, ids, options, odds, reason: 'combine adjacent forces to open an objective route without exhausting supply' };
      }).filter(battle => battle.odds >= (unit.side === 'german' ? 1.15 : 1.35))
        .sort((a, b) => (b.odds + (state.map[base.key(b.defender.x, b.defender.y)].value || 0) * 0.2) -
          (a.odds + (state.map[base.key(a.defender.x, a.defender.y)].value || 0) * 0.2));
      return targets[0] || null;
    }
    function preview(state, ids, defenderId, options = {}) {
      const defender = state.units.find(unit => unit.id === defenderId), eligible = eligibleAttackers(state, defender);
      if (!ids.length || ids.length > 6 || new Set(ids).size !== ids.length ||
        ids.some(id => !eligible.some(unit => unit.id === id))) throw new Error('Choose eligible adjacent attackers.');
      const strategy = options.strategy || 'assault', defenseStrategy = options.defenseStrategy ||
        (defender.side === state.playerSide ? state.defenseStrategy : 'flexible'), artillery = options.artillery ?? 0;
      if (!Object.hasOwn(ATTACK, strategy) || !Object.hasOwn(DEFENSE, defenseStrategy) ||
        !Number.isInteger(artillery) || artillery < 0 || artillery > Math.min(ids.length, state.artillery[state.phase])) {
        throw new Error('Invalid combat strategy or artillery allocation.');
      }
      const attackers = ids.map(id => state.units.find(unit => unit.id === id));
      const attackTotal = attackers.reduce((sum, unit) => sum + strength(state, unit), 0) + artillery * 10;
      const defensiveArtillery = Math.min(2, state.artillery[defender.side]);
      const defenseTotal = strength(state, defender, true) + defensiveArtillery * 10;
      const odds = attackTotal / Math.max(1, defenseTotal);
      const tile = state.map[base.key(defender.x, defender.y)];
      return { odds, detail: `${attackers.map(unit => `${unit.name}: ${unit.hp} CP, readiness x${(1 + unit.readiness / 10).toFixed(1)}, ${localSupply(state, unit)} supply x${modifier(localSupply(state, unit))}`).join('; ')}. Attack total ${attackTotal.toFixed(1)} (artillery +${artillery * 10}). ${defender.name}: ${defender.hp} CP, readiness x${(1 + defender.readiness / 10).toFixed(1)}, ${localSupply(state, defender)} supply x${modifier(localSupply(state, defender))}, ${tile.terrain} defense x${terrainDefense(state, tile)}. Defense total ${defenseTotal.toFixed(1)} (artillery +${defensiveArtillery * 10}).`,
        risk: `${odds < 1 ? 'High risk' : odds < 1.7 ? 'Contested' : 'Favorable'} nominal odds before random general-supply failures. Supply delivery chance: attacker ${Math.round(state.generalSupply[state.phase] / 120 * 100)}%, defender ${Math.round(state.generalSupply[defender.side] / 120 * 100)}%. ${strategyLabel(strategy)} / ${strategyLabel(defenseStrategy)} commitment x${(ATTACK[strategy] + DEFENSE[defenseStrategy]).toFixed(2)} affects losses and consumption, not these nominal odds. Weather ${state.weather}: Allied air losses may apply. Retreats, advances, order changes and low-strength elimination remain uncertain. This is an adapted model, not the original hidden combat table.` };
    }
    function reinforce(state) {
      state.arrivals = [];
      for (const entry of schedule(state).filter(entry => entry.turn <= state.turn && !state.units.some(unit => unit.id === entry.id))) {
        const desired = Object.values(state.map).find(tile => tile.col === entry.col && tile.row === entry.row);
        let tile = !base.occupant(state, desired) ? desired : null;
        if (!tile && state.pending.includes(entry.id)) {
          tile = Object.values(state.map).filter(tile => entry.side === 'german' ? tile.col >= 12 : tile.col <= 4)
            .sort((a, b) => base.distance(a, desired) - base.distance(b, desired)).find(tile => !base.occupant(state, tile));
        }
        if (!tile) {
          if (!state.pending.includes(entry.id)) state.pending.push(entry.id);
          log(state, `${entry.name}: entry blocked; arrival delayed until a later reinforcement phase.`);
          base.reportOperation(state, entry.side, `${entry.name} could not arrive because its entry hex is occupied. It will try again in a later reinforcement phase.`);
          continue;
        }
        state.units.push({ id: entry.id, name: entry.name, side: entry.side, type: entry.type, x: tile.x, y: tile.y,
          hp: entry.hp, maxHp: entry.hp, attack: entry.type === 'armor' ? 8 : 5, defense: entry.type === 'armor' ? 6 : 5,
          mobility: entry.type === 'armor' ? 24 : 20, mp: 0, fired: false, defended: false, exited: false,
          supplied: true, supplyState: 'supplied', readiness: 6 });
        state.pending = state.pending.filter(id => id !== entry.id);
        state.arrivals.push(entry.id); capture(state, state.units[state.units.length - 1]); log(state, `${entry.name} arrives.`);
        base.reportOperation(state, entry.side, `${entry.name} arrived with ${entry.hp} combat points at column ${tile.col + 1}, row ${tile.row + 1}. Look for NEW on the map.`);
      }
    }
    function finishDay(state) {
      reinforce(state);
      for (const unit of state.units.filter(living)) {
        const tile = state.map[base.key(unit.x, unit.y)];
        if (unit.side === 'german' && tile.col === 0 && !enemyNear(state, unit, unit.side)) {
          unit.exited = true; unit.mp = 0; state.exitPoints += Math.ceil(unit.hp / 10);
          log(state, `${unit.name} exits west; victory points awarded.`);
          base.reportOperation(state, unit.side, `${unit.name} broke through and exited the western map edge, earning ${Math.ceil(unit.hp / 10)} victory points.`);
        }
      }
      state.alliedBreakthrough = state.units.filter(unit => living(unit) && unit.side === 'american' &&
        state.map[base.key(unit.x, unit.y)].col === 14).length * 2;
      score(state);
      base.summarizeDay(state);
      const germanSurvives = state.units.some(unit => living(unit) && unit.side === 'german') || state.exitPoints > 0;
      if (state.turn >= 12 || !state.units.some(unit => living(unit) && unit.side === 'american') ||
        (!state.units.some(unit => living(unit) && unit.side === 'german') && state.exitPoints === 0)) {
        state.winner = germanSurvives && (state.score >= 8 || !state.units.some(unit => living(unit) && unit.side === 'american')) ? 'german' : 'american';
        state.phase = 'finished'; log(state, `${state.winner.toUpperCase()} VICTORY / adaptation score ${state.score}.`);
        return;
      }
      state.turn++;
      state.weather = state.turn < 4 ? 'Overcast' : state.turn < 8 ?
        (random(state) < 0.6 ? 'Cloudy' : 'Scattered') : (random(state) < 0.8 ? 'Clear' : 'Scattered');
      state.generalSupply.german = Math.min(120, state.generalSupply.german + 50);
      state.generalSupply.american = Math.min(120, state.generalSupply.american + 60);
      state.artillery = { german: 6, american: 4 };
      if (['Clear', 'Scattered'].includes(state.weather)) {
        for (const unit of state.units.filter(unit => living(unit) && unit.side === 'american')) {
          if (random(state) < 0.3 && localSupply(state, unit) !== 'supplied') {
            unit.airdropTurn = state.turn;
            log(state, `${unit.name}: supply airdrop for this day.`);
          }
        }
      }
      log(state, `${15 + state.turn} DEC / ${state.weather.toUpperCase()} / new day.`);
      state.dayStart = base.daySnapshot(state);
    }
    function endTurn(state, side) {
      if (state.winner || state.phase !== side) return false;
      if (state.orderPhase === 'movement') { state.orderPhase = 'combat'; return true; }
      if (side === 'german') { computerTurn(state, 'american'); finishDay(state); }
      else finishDay(state);
      if (!state.winner) {
        beginSide(state, 'german');
        if (side === 'american') { computerTurn(state, 'german'); beginSide(state, 'american'); }
      }
      return true;
    }
    function deserialize(parsed) {
      const integer = (value, min, max) => Number.isInteger(value) && value >= min && value <= max;
      if (parsed.ruleset !== 'manual' || !integer(parsed.turn, 1, 12) || !integer(parsed.seed, 0, 4294967295) ||
        !['german', 'american'].includes(parsed.playerSide) || !['movement', 'combat'].includes(parsed.orderPhase) ||
        !['german', 'american', 'finished'].includes(parsed.phase) || !['Overcast', 'Cloudy', 'Scattered', 'Clear'].includes(parsed.weather) ||
        ![null, 'german', 'american'].includes(parsed.winner) || (parsed.phase === 'finished') !== Boolean(parsed.winner) ||
        (parsed.playMode !== undefined && !['human', 'computer'].includes(parsed.playMode)) ||
        (parsed.playMode === 'computer' && parsed.playerSide !== 'german') ||
        (!parsed.winner && parsed.phase !== parsed.playerSide && !(parsed.playMode === 'computer' && parsed.computerProgress)) || !integer(parsed.fuel, 0, 120) ||
        (parsed.referenceVersion !== undefined && parsed.referenceVersion !== 1) ||
        !integer(parsed.exitPoints, 0, 1000) || !integer(parsed.alliedBreakthrough, 0, 100) ||
        !Object.hasOwn(DEFENSE, parsed.defenseStrategy) || !parsed.generalSupply || !parsed.artillery ||
        !['german', 'american'].every(side => integer(parsed.generalSupply[side], 0, 120) &&
          integer(parsed.artillery[side], 0, side === 'german' ? 6 : 4)) ||
        !Array.isArray(parsed.units) || parsed.units.length < 14 || parsed.units.length > 21 ||
        !Array.isArray(parsed.log) || parsed.log.length > 60 || parsed.log.some(text => typeof text !== 'string' || text.length > 250)) {
        throw new Error('Invalid manual-mode campaign data.');
      }
      const template = base.createGame(parsed.seed);
      const known = [...template.units, ...schedule(parsed)];
      const ids = new Set(), occupied = new Set();
      const units = parsed.units.map(unit => {
        if (!unit || typeof unit !== 'object') throw new Error('Invalid manual-mode unit data.');
        const expected = known.find(entry => entry.id === unit.id);
        if (!expected || ids.has(unit.id) || unit.side !== expected.side || unit.type !== expected.type ||
          typeof unit.name !== 'string' || unit.name.length > 50 || !integer(unit.x, 0, 29) || !integer(unit.y, 0, 11) ||
          !template.map[base.key(unit.x, unit.y)] || !integer(unit.hp, 0, 120) || !integer(unit.maxHp, 1, 120) ||
          unit.hp > unit.maxHp || !integer(unit.mobility, 1, 24) || !integer(unit.mp, 0, unit.mobility) ||
          !integer(unit.attack, 1, 12) || !integer(unit.defense, 1, 12) || !integer(unit.readiness, 0, 7) ||
          !['fired', 'defended', 'exited', 'supplied'].every(field => typeof unit[field] === 'boolean') ||
          !['supplied', 'limited', 'isolated'].includes(unit.supplyState) ||
          (unit.airdropTurn !== undefined && !integer(unit.airdropTurn, 1, parsed.turn)) ||
          (living(unit) && occupied.has(base.key(unit.x, unit.y)))) throw new Error('Invalid manual-mode unit data.');
        ids.add(unit.id); if (living(unit)) occupied.add(base.key(unit.x, unit.y));
        return { id: unit.id, name: unit.name, side: unit.side, type: unit.type, x: unit.x, y: unit.y, hp: unit.hp, maxHp: unit.maxHp,
          mobility: unit.mobility, mp: unit.mp, attack: unit.attack, defense: unit.defense, readiness: unit.readiness,
          fired: unit.fired, defended: unit.defended, exited: unit.exited, supplied: unit.supplied, supplyState: unit.supplyState,
          ...(unit.airdropTurn === undefined ? {} : { airdropTurn: unit.airdropTurn }) };
      });
      if (template.units.some(unit => !ids.has(unit.id)) || !Array.isArray(parsed.arrivals) || !Array.isArray(parsed.pending) ||
        parsed.arrivals.length > 7 || parsed.pending.length > 7 || new Set(parsed.arrivals).size !== parsed.arrivals.length ||
        new Set(parsed.pending).size !== parsed.pending.length ||
        parsed.arrivals.some(id => !schedule(parsed).some(entry => entry.id === id) || !ids.has(id)) ||
        parsed.pending.some(id => !schedule(parsed).some(entry => entry.id === id) || ids.has(id))) throw new Error('Invalid reinforcement data.');
      for (const tile of Object.values(template.map).filter(tile => tile.value)) {
        const saved = parsed.map && parsed.map[base.key(tile.x, tile.y)];
        if (!saved || !['german', 'american'].includes(saved.owner)) throw new Error('Invalid objective data.');
        tile.owner = saved.owner;
      }
      Object.assign(template, { version: 2, ruleset: 'manual', playerSide: parsed.playerSide, turn: parsed.turn, phase: parsed.phase,
        orderPhase: parsed.orderPhase, weather: parsed.weather, seed: parsed.seed, winner: parsed.winner,
        fuel: parsed.fuel, exitPoints: parsed.exitPoints, alliedBreakthrough: parsed.alliedBreakthrough,
        generalSupply: { german: parsed.generalSupply.german, american: parsed.generalSupply.american },
        artillery: { german: parsed.artillery.german, american: parsed.artillery.american }, defenseStrategy: parsed.defenseStrategy,
        units, log: [...parsed.log], arrivals: [...parsed.arrivals], pending: [...parsed.pending] });
      if (parsed.playMode !== undefined) template.playMode = parsed.playMode;
      if (parsed.referenceVersion !== undefined) template.referenceVersion = parsed.referenceVersion;
      score(template); return base.restoreMetadata(parsed, template);
    }
    return { initialize, beginSide, localSupply, refreshSupply, reachable, move, combatOdds, eligibleAttackers,
      resolveAttack, computerTurn, endTurn, deserialize, reinforce, finishDay, COST, SCHEDULE, preview, planBattle, terrainDefense, strategyLabel };
  }
  if (typeof module === 'object') module.exports = createManualRules;
  else root.createManualRules = createManualRules;
})(typeof globalThis !== 'undefined' ? globalThis : this);
