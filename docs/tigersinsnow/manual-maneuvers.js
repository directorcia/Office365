(function (root) {
  'use strict';
  function createManualManeuvers(base, rules) {
    const living = unit => unit && unit.hp > 0 && !unit.exited;
    const human = (state, unit) => state.playMode === 'hotseat' ||
      (state.playMode !== 'computer' && unit.side === state.playerSide);
    function current(state) {
      return state.pendingBattle?.tasks[0] || null;
    }
    function choices(state) {
      const task = current(state);
      if (!task) return [];
      const unit = state.units.find(unit => unit.id === task.id);
      if (!living(unit)) return [];
      return base.aidCard.directions.map(([dx,dy], i) => {
        const tile = state.map[base.key(unit.x + dx, unit.y + dy)];
        if (!tile || base.occupant(state, tile) || task.path.includes(base.key(tile.x, tile.y))) return null;
        if (task.kind === 'advance' && !task.path.length &&
          base.key(tile.x, tile.y) !== base.key(task.first.x, task.first.y)) return null;
        return { direction: i + 1, tile };
      }).filter(Boolean);
    }
    function taskFor(id, kind, count, first) {
      return { id, kind, remaining: count, path: [], ...(first ? { first } : {}) };
    }
    function finishTask(state, stopped) {
      const battle = state.pendingBattle, task = battle.tasks.shift();
      const unit = state.units.find(unit => unit.id === task.id);
      if (task.kind === 'retreat' && stopped && task.remaining && living(unit)) {
        rules.lose(state, unit, task.remaining * (2 + Math.floor(rules.random(state) * 4)));
      }
      const path = { id: task.id, path: [...task.path] };
      if (task.kind === 'retreat') {
        if (task.id === battle.defenderId) battle.retreatPath = path.path;
        else battle.attackerRetreats.push(path);
      } else if (task.path.length) battle.advancePaths.push(path);
      battle.events.push(`${unit.name}: ${task.kind} ${task.path.length} hex(es)${stopped && task.remaining && task.kind === 'retreat' ? '; incomplete retreat caused additional losses' : ''}.`);
    }
    function prepareAdvances(state) {
      const battle = state.pendingBattle;
      battle.stage = 'advances';
      if (rules.supplement(state)) {
        battle.secondCombatIds = [];
        const candidates = battle.ids.filter(id => {
          const unit = state.units.find(unit => unit.id === id);
          return living(unit) && unit.combatCount === 1;
        });
        if (candidates.length && battle.advanceLimit >= 2 && !base.occupant(state, battle.original)) {
          const eligible = battle.advanceLimit >= 3 || rules.random(state) < 0.25;
          if (eligible) battle.secondCombatIds = candidates;
        }
        if (battle.secondCombatIds.length) {
          const message = `Units eligible for second combat: ${battle.secondCombatIds.map(id =>
            state.units.find(unit => unit.id === id).name).join(', ')}.`;
          battle.events.push(message);
          for (const id of battle.secondCombatIds) {
            const notice = `${state.units.find(unit => unit.id === id).name}: eligible for second combat.`;
            rules.log(state, notice); base.reportOperation(state, state.phase, notice);
          }
        }
      }
      if (battle.advanceLimit > 0) {
        for (const id of battle.ids) {
          const unit = state.units.find(unit => unit.id === id);
          if (living(unit) && base.distance(unit, battle.original) === 1) {
            battle.tasks.push(taskFor(id, 'advance', battle.advanceLimit, battle.original));
          }
        }
      } else if (battle.attackerRetreats.length) {
        const defender = state.units.find(unit => unit.id === battle.defenderId);
        const vacated = battle.attackerPositions.find(tile => !base.occupant(state, tile) && base.distance(defender, tile) === 1);
        if (living(defender) && vacated) battle.tasks.push(taskFor(defender.id, 'advance', 1, vacated));
      }
    }
    function complete(state) {
      const battle = state.pendingBattle;
      const units = battle.before.map(snapshot => {
        const unit = state.units.find(unit => unit.id === snapshot.id);
        return { ...snapshot, remaining: unit.hp, loss: snapshot.hp - unit.hp };
      });
      if (rules.supplement(state)) {
        for (const id of battle.secondCombatIds || []) {
          const unit = state.units.find(unit => unit.id === id);
          if (living(unit)) { unit.fired = false; battle.events.push(`${unit.name} may attack again.`); }
        }
        const progress = state.aidFlow || state.computerProgress;
        if (progress?.stage === 'combat' && battle.secondCombatIds?.length) {
          const units = state.units.filter(unit => unit.side === state.phase);
          const next = units.findIndex(unit => battle.secondCombatIds.includes(unit.id) && living(unit));
          if (next >= 0) progress.index = Math.min(progress.index, next);
        }
      } else if (battle.advanceLimit >= 3) {
        for (const id of battle.ids) {
          const unit = state.units.find(unit => unit.id === id);
          if (living(unit) && rules.random(state) < 0.5) {
            unit.fired = false; battle.events.push(`${unit.name} may attack again.`);
          }
        }
      }
      rules.refreshSupply(state); rules.score(state);
      rules.log(state, `${battle.ids.map(id => state.units.find(unit => unit.id === id).name).join(' + ')} > ${state.units.find(unit => unit.id === battle.defenderId).name}: ${battle.odds.toFixed(1)}:1 / ${battle.attackStrategy} vs ${battle.defenseStrategy}.`);
      for (const event of battle.events) rules.log(state, event);
      for (const unit of units) {
        base.reportOperation(state, unit.side, `${unit.name}: battle complete, ${unit.loss} CP lost, ${unit.remaining} remaining. ${battle.events.filter(event => event.startsWith(unit.name)).join(' ')}`);
      }
      const result = { odds: battle.odds, attackStrategy: battle.attackStrategy, defenseStrategy: battle.defenseStrategy,
        artillery: battle.artillery, defensiveArtillery: battle.defensiveArtillery, units, events: battle.events,
        retreatPath: battle.retreatPath, attackerRetreats: battle.attackerRetreats, advancePaths: battle.advancePaths,
        ...(rules.supplement(state) ? { secondCombatIds: [...(battle.secondCombatIds || [])] } : {}) };
      delete state.pendingBattle;
      return result;
    }
    function step(state, direction) {
      const task = current(state), unit = state.units.find(unit => unit.id === task.id);
      if (direction === 'finish') { finishTask(state, true); return; }
      const choice = choices(state).find(choice => choice.direction === direction);
      if (!choice) throw new Error('That maneuver direction is blocked or does not enter the required vacated hex.');
      unit.x = choice.tile.x; unit.y = choice.tile.y; unit.readiness = Math.max(0, unit.readiness - 2);
      task.path.push(base.key(unit.x, unit.y)); task.remaining--;
      rules.capture(state, unit);
      const enemyZone = rules.enemyNear(state, unit, unit.side);
      if (task.kind === 'retreat' && enemyZone) rules.lose(state, unit, 1 + Math.floor(rules.random(state) * 6));
      if (!task.remaining || !living(unit) || (task.kind === 'advance' && enemyZone)) finishTask(state, false);
    }
    function process(state) {
      for (let i = 0; i < 150; i++) {
        const battle = state.pendingBattle;
        if (!battle.tasks.length) {
          if (battle.stage === 'retreats') { prepareAdvances(state); continue; }
          return complete(state);
        }
        const task = current(state), unit = state.units.find(unit => unit.id === task.id);
        if (!living(unit) || (task.kind === 'retreat' && rules.localSupply(state, unit) === 'isolated') || !choices(state).length) {
          finishTask(state, true); continue;
        }
        if (human(state, unit)) {
          const units = battle.before.map(snapshot => {
            const now = state.units.find(unit => unit.id === snapshot.id);
            return { ...snapshot, remaining: now.hp, loss: snapshot.hp - now.hp };
          });
          return { ...battle, units, pending: true,
            events: [...battle.events, `${unit.name}: choose ${task.kind} directions on the maneuver panel.`] };
        }
        const enemies = state.units.filter(enemy => living(enemy) && enemy.side !== unit.side);
        const ranked = choices(state).sort((a,b) => {
          const rank = tile => task.kind === 'retreat' ?
            Math.min(...enemies.map(enemy => base.distance(tile, enemy)), 100) * 10 + (unit.side === 'german' ? tile.col : -tile.col) :
            unit.side === 'german' ? -tile.col : tile.col;
          return rank(b.tile) - rank(a.tile);
        });
        step(state, ranked[0].direction);
      }
      throw new Error('Battle maneuver processing exceeded its step limit.');
    }
    function start(state, battle) {
      state.pendingBattle = { ...battle, stage: 'retreats', tasks: [], retreatPath: [], attackerRetreats: [], advancePaths: [] };
      if (battle.retreatCount > 0) state.pendingBattle.tasks.push(taskFor(battle.defenderId, 'retreat', battle.retreatCount));
      else if (battle.attackerMustRetreat) {
        for (const id of battle.ids) state.pendingBattle.tasks.push(taskFor(id, 'retreat', 1));
      }
      return process(state);
    }
    function choose(state, direction) {
      if (!current(state)) throw new Error('There is no pending battle maneuver.');
      const unit = state.units.find(unit => unit.id === current(state).id);
      if (!human(state, unit)) throw new Error('This maneuver belongs to the computer.');
      step(state, direction); return process(state);
    }
    function validate(saved, state) {
      if (saved === undefined) return;
      const unitIds = new Set(state.units.map(unit => unit.id));
      const tile = value => value && Number.isInteger(value.x) && Number.isInteger(value.y) && state.map[base.key(value.x, value.y)];
      const ids = values => Array.isArray(values) && values.length >= 1 && values.length <= 6 &&
        new Set(values).size === values.length && values.every(id => unitIds.has(id));
      const count = value => Number.isInteger(value) && value >= 0 && value <= 6;
      const paths = values => Array.isArray(values) && values.length <= 6 && values.every(value =>
        value && unitIds.has(value.id) && Array.isArray(value.path) && value.path.length <= 6 && value.path.every(key => state.map[key]));
      if (!saved || !ids(saved.ids) || !unitIds.has(saved.defenderId) || !tile(saved.original) ||
        !Number.isFinite(saved.odds) || saved.odds < 0 || saved.odds > 10000 ||
        !['retreats','advances'].includes(saved.stage) || !count(saved.retreatCount) || !count(saved.advanceLimit) ||
        typeof saved.attackerMustRetreat !== 'boolean' || !Array.isArray(saved.tasks) || saved.tasks.length < 1 || saved.tasks.length > 6 ||
        saved.tasks.some(task => !task || !unitIds.has(task.id) || !['retreat','advance'].includes(task.kind) ||
          !count(task.remaining) || task.remaining < 1 || !Array.isArray(task.path) || task.path.length > 6 ||
          task.path.some(key => !state.map[key]) || (task.kind === 'advance' && !tile(task.first))) ||
        !Array.isArray(saved.before) || saved.before.length !== saved.ids.length + 1 ||
        saved.before.some(value => !value || !unitIds.has(value.id) || !Number.isInteger(value.hp) || value.hp < 1 || value.hp > 132 ||
          value.name !== state.units.find(unit => unit.id === value.id).name || value.side !== state.units.find(unit => unit.id === value.id).side) ||
        !Array.isArray(saved.attackerPositions) || saved.attackerPositions.length !== saved.ids.length ||
        saved.attackerPositions.some(value => !value || !saved.ids.includes(value.id) || !tile(value)) ||
        !Array.isArray(saved.events) || saved.events.length > 60 || saved.events.some(value => typeof value !== 'string' || value.length > 500) ||
        !Array.isArray(saved.retreatPath) || saved.retreatPath.length > 6 || saved.retreatPath.some(key => !state.map[key]) ||
        !paths(saved.attackerRetreats) || !paths(saved.advancePaths) ||
        !['recon','probe','assault','major'].includes(saved.attackStrategy) ||
        !['counter','hold','flexible','delay','withdrawal'].includes(saved.defenseStrategy) ||
        !Number.isInteger(saved.artillery) || saved.artillery < 0 || saved.artillery > 6 ||
        !Number.isInteger(saved.defensiveArtillery) || saved.defensiveArtillery < 0 || saved.defensiveArtillery > 6) {
        throw new Error('Invalid saved battle maneuver.');
      }
      if (rules.supplement(state) && saved.stage === 'advances' &&
        (!Array.isArray(saved.secondCombatIds) || new Set(saved.secondCombatIds).size !== saved.secondCombatIds.length ||
          saved.secondCombatIds.some(id => !saved.ids.includes(id) || state.units.find(unit => unit.id === id).combatCount !== 1) ||
          (saved.advanceLimit < 2 && saved.secondCombatIds.length))) {
        throw new Error('Invalid saved second-combat eligibility.');
      }
      state.pendingBattle = structuredClone(saved);
      if (!choices(state).length || !human(state, state.units.find(unit => unit.id === current(state).id))) {
        throw new Error('Saved maneuver is not awaiting a legal human decision.');
      }
    }
    return { start, choose, current, choices, validate };
  }
  if (typeof module === 'object') module.exports = createManualManeuvers;
  else root.createManualManeuvers = createManualManeuvers;
})(typeof globalThis !== 'undefined' ? globalThis : this);
