(function (root) {
  'use strict';
  const german = [
    ['277th VG DIV',55],['272nd VG DIV',50],['326th VG DIV',50],['12th VG DIV',55],
    ['560th VG DIV',50],['18th VG DIV',35],['62nd VG DIV',50],['26th VG DIV',65],
    ['352nd VG DIV',65],['276th VG DIV',55],['212th VG DIV',65],['3rd FJ DIV',70],['5th FJ DIV',70],
    ['116th PZ DIV',100,'armor'],['2nd PZ DIV',100,'armor'],['LEHR PZ DIV',100,'armor'],
    ['1st SS PZ DIV',115,'armor'],
    ['12th SS PZ DIV',115,'armor',16,13],['3rd PZGR DIV',75,'armor',16,12],
    ['150th PZ BRIG',45,'armor',16,11],['9th SS PZ DIV',115,'armor',18,14],
    ['2nd SS PZ DIV',115,'armor',18,13],['FBE PZ BRIG',50,'armor',18,12],
    ['79th VG DIV',60,'infantry',22,16],['15th PZGR DIV',75,'armor',22,14],
    ['9th PZ DIV',100,'armor',22,15],['FGR PZ BRIG',45,'armor',22,13],
    ['167th VG DIV',55,'infantry',24,15]
  ];
  const allied = [
    ['2nd INF DIV',80],['99th INF DIV',75],['4th INF DIV',75],['106th INF DIV',55],
    ['112/28 INF REG',20],['110/28 INF REG',20],['109/28 INF REG',30],['9th INF DIV',75],
    ['CCA/9 ARM REG',45,'armor'],['CCB/9 ARM REG',35,'armor'],['CCR/9 ARM REG',35,'armor'],
    ['14th CALV REG',30,'armor'],
    ['CCA/10 ARM REG',40,'armor',16,18],['CCB/10 ARM REG',40,'armor',16,20],
    ['CCR/10 ARM REG',40,'armor',16,17],['CCA/7 ARM REG',40,'armor',16,10],
    ['CCB/7 ARM REG',40,'armor',16,9],['CCR/7 ARM REG',40,'armor',16,8],
    ['82nd AIRB DIV',70,'infantry',16,28],['101st AIRB DIV',70,'infantry',16,24],
    ['1st INF DIV',75,'infantry',17,10],['30th INF DIV',70,'infantry',19,8],
    ['3rd ARM DIV',100,'armor',20,6],['4th ARM DIV',100,'armor',20,21],
    ['43rd INF DIV',65,'infantry',20,7,'british'],['29th ARM BRIG',45,'armor',21,26,'british'],
    ['80th INF DIV',70,'infantry',21,20],['5th INF DIV',70,'infantry',21,19],
    ['84th INF DIV',70,'infantry',21,5],['26th INF DIV',55,'infantry',21,22],
    ['2nd ARM DIV',100,'armor',22,2],['GDS ARM DIV',100,'armor',22,1,'british'],
    ['53rd INF DIV',65,'infantry',22,27,'british'],['33rd ARM BRIG',35,'armor',22,3,'british'],
    ['75th INF DIV',70,'infantry',23,25],['51st INF DIV',65,'infantry',24,4,'british'],
    ['35th INF DIV',70,'infantry',25,23],['83rd INF DIV',70,'infantry',25,6]
  ];
  const roster = [...german.map((row, i) => ['german', `card-g${i + 1}`, row]),
    ...allied.map((row, i) => ['american', `card-a${i + 1}`, row])]
    .map(([side, id, [name, hp, type = 'infantry', day, entry, nationality]]) =>
      ({ id, name, side, hp, type, turn: day === undefined ? 0 : day - 15, entry,
        nationality: nationality || (side === 'german' ? 'german' : 'us') }));
  const abbreviations = {
    INF: 'Infantry', ARM: 'Armor (mechanized)', AIRB: 'Airborne', CALV: 'Armored cavalry (mechanized)',
    FJ: 'Fallschirmjager (parachute)', PZ: 'Panzer (mechanized armor)',
    PZGR: 'Panzer grenadier (mechanized armored infantry)', VG: 'Volksgrenadier infantry',
    FGR: 'Fuehrer Grenadier', FBE: 'Fuehrer Begleit', REG: 'Regiment', DIV: 'Division',
    BRIG: 'Brigade', BATT: 'Battalion', CCA: 'Combat Command A', CCB: 'Combat Command B',
    CCR: 'Combat Command R', 'XXX/28': 'A regiment of the 28th Division', GDS: 'Guards'
  };
  const directions = [[1,-1],[2,0],[1,1],[-1,1],[-2,0],[-1,-1]];
  const levels = [
    ['american','Strategic',50,50],['american','Operational',299,174],['american','Marginal',699,324],
    ['german','Marginal',949,424],['german','Operational',1199,699],['german','Strategic',Infinity,Infinity]
  ];
  const active = state => state.aidCardVersion === 1;
  function options(value = {}) {
    const ratings = value.ratings || { german: 5, american: 5 };
    if (!['campaign', 'short'].includes(value.scenario || 'campaign') ||
      !['german','american'].every(side => Number.isInteger(ratings[side]) && ratings[side] >= 1 && ratings[side] <= 9)) {
      throw new Error('Choose a valid scenario and whole-number ratings from 1 to 9.');
    }
    if (value.display !== undefined && (!value.display || typeof value.display.color !== 'boolean' || typeof value.display.hexes !== 'boolean')) {
      throw new Error('Choose valid color and hex-grid display settings.');
    }
    return { scenario: value.scenario || 'campaign', ratings: { german: ratings.german, american: ratings.american },
      display: value.display ? { color: value.display.color, hexes: value.display.hexes } : { color: true, hexes: true } };
  }
  const strength = (entry, state) => entry.hp + 3 * (state.ratings[entry.side] - 5);
  const lastTurn = state => active(state) && state.scenario === 'short' ? 7 : 12;
  const computerGerman = state => state.playMode === 'computer' ||
    (state.playMode !== 'hotseat' && state.playerSide === 'american');
  const multiplier = state => computerGerman(state) ? 2 : 1;
  function edgeAward(tile, side) {
    if (side === 'german') {
      if ((tile.row === 10 && tile.col <= 13) || (tile.col === 0 && tile.row >= 8)) return 3;
      if (tile.row === 0 && tile.col >= 10 && tile.col <= 13) return 10;
      return 0;
    }
    return tile.col === 18 && tile.row >= 4 && tile.row <= 6 ? 20 : 0;
  }
  function standing(state) {
    const points = state.victoryPoints.german - state.victoryPoints.american;
    const level = levels.find(row => points <= row[state.scenario === 'short' ? 3 : 2]);
    const target = state.scenario === 'short' ? 325 : 700;
    return { points, target, leader: level[0], level: level[1],
      status: `${level[0] === 'american' ? 'ALLIED' : 'GERMAN'} ${level[1].toUpperCase()} ${state.winner ? 'VICTORY' : 'ON COURSE'}`,
      detail: `${state.victoryPoints.german} German VP − ${state.victoryPoints.american} Allied VP = ${points}. ${state.scenario === 'short' ? 'Bastogne–St. Vith, 16–22 December' : 'Campaign, 16–27 December'}.`,
      provisional: !state.winner };
  }
  function assessDay(state, base) {
    const factor = multiplier(state), changes = { german: 0, american: 0 };
    for (const unit of state.units.filter(unit => unit.hp > 0 && !unit.exited)) {
      const tile = state.map[base.key(unit.x, unit.y)];
      if (unit.side === 'german') {
        changes.german += edgeAward(tile, 'german') * factor;
        if (tile.sourcePoints) changes.german += tile.sourcePoints * factor;
        if (tile.exit && !base.neighbors(state, tile).some(next => {
          const enemy = base.occupant(state, next);
          return enemy && enemy.side !== unit.side;
        })) {
          const points = unit.hp * (16 - state.turn) / 4;
          changes.german += points; state.exitPoints += points; unit.exited = true; unit.mp = 0;
          base.reportOperation(state, unit.side, `${unit.name} exited at ${unit.hp} CP × (16 − ${state.turn}) / 4, earning ${points} VP.`);
        }
      } else changes.american += edgeAward(tile, 'american');
    }
    state.victoryPoints.german += changes.german;
    state.victoryPoints.american += changes.american;
    state.alliedBreakthrough = changes.american;
    state.score = standing(state).points;
    return changes;
  }
  const api = { roster, abbreviations, directions, levels, active, options, strength, lastTurn,
    multiplier, standing, assessDay, edgeAward };
  if (typeof module === 'object') module.exports = api;
  else root.AidCard = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
